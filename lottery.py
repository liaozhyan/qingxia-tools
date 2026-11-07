# -*- coding: utf-8 -*-
"""现场抽签排序（原生窗口版）。

用于企业宣讲比赛、路演、评审答辩等场合，公开、可审计地决定出场顺序。
从独立 Flask 项目（抽签排序/）合并而来，改为同进程 Tkinter 原生窗口：

  * LotteryAdmin  —— 管理窗口：设置 / 名单（Excel·CSV·粘贴导入）/
                     抽签 / 结果 / 历史记录 / 校验 / 导出 / 背景音乐
  * LotteryScreen —— 大屏投屏窗口：卡片网格 + 轮播 → 逐个定格 →
                     撒花 + 音效，投到扩展屏全屏

设计要点（与 Flask 版一致）：
1. 随机数使用 random.Random(seed)，seed 由 secrets.randbits(64) 生成并
   全程留痕，任何一次抽签结果都可以用 seed 复现，回应"是否内定"质疑。
2. 每次抽签落盘 data/draws/<id>.json，含原始名单快照，支持事后复现校验。
3. 数据沿用旧版 抽签排序/data/state.json（开发目录存在时），
   旧参与者名单与历史记录直接保留。
4. 零界面三方依赖：音效用 winsound（系统蜂鸣），背景音乐走 MCI
   （winmm.dll），Excel 读写用 openpyxl（惰性导入，缺失时提示）。
"""

import csv
import io
import json
import math
import os
import random
import secrets
import sys
import threading
import time
import uuid
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

import ui_theme
from ui_theme import PALETTE, FlatButton, GradientBand, font, round_rect

ALGORITHM = 'random.Random(seed) + Fisher-Yates shuffle'
ALLOWED_MUSIC_EXT = ('.mp3', '.wav', '.ogg', '.m4a', '.flac', '.aac', '.mid', '.wma')

ACCENT = ui_theme.ACCENTS.get('lottery', '#65a30d')
ACCENT_DARK = '#3f6212'

# 大屏配色（深色底）
SC_BG = '#0b1530'
SC_BG2 = '#13224a'
SC_CARD = '#101c3a'
SC_CARD_LOCK = '#152f1c'
SC_LINE = '#24365e'
SC_LINE_LOCK = '#7cb342'
SC_TEXT = '#f1f5ff'
SC_MUTED = '#8ea0c8'
SC_FLASH = '#ffd54f'

ROLL_MS = 1700      # 快速轮播时长
CYCLE_MS = 65       # 轮播帧间隔
# 滚动阶段占位用"流光彗尾"：尾迹圆点自左向右扫过并渐弱，
# 尺寸按卡宽约束、始终装进卡片，且不含姓名——名字只在定格时出现
DOTS_N = 5
DOT_OFF = '#2c3a5c'       # 熄灭
DOT_HEAD = '#eaf6ff'       # 亮头
DOT_T1 = '#8fd0ff'         # 尾迹一
DOT_T2 = '#5b7bb7'         # 尾迹二
DOT_COLORS = (DOT_HEAD, DOT_T1, DOT_T2)   # 按到亮头的距离取色
# 可选的抽签动画样式（滚动阶段占位效果）
STYLE_DOTS = '彗尾流光'        # 圆点彗尾扫过（默认）
STYLE_RING = '光环脉冲'        # 双环由外向内脉动收拢
STYLE_DIGITS = '数字滚轮'      # 大数字快速翻转
STYLE_OPTIONS = (STYLE_DOTS, STYLE_RING, STYLE_DIGITS)
STYLE_DEF_DEFAULT = STYLE_DOTS
# 自动揭晓预算：Tk 画布每卡重绘实测有 ~230ms 开销（2560x1440 大画布），
# 把预算压到 4.5s，实际 21 卡总揭晓约 9~10s，与网页版 9s 节奏一致
AUTO_TOTAL_MS = 4500
# 抽签速度档位（倍率越大越快）：同时缩放轮播时长、帧间隔、揭晓总预算
SPEED_PRESETS = {'慢': 0.6, '正常': 1.0, '快': 1.6, '极速': 2.5}


# ---------------------------------------------------------------------------
# 数据目录（开发态沿用 抽签排序/data，保留旧数据；打包后用 EXE 旁的 data）
# ---------------------------------------------------------------------------
def _base_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


_BASE = _base_dir()
_DEV_DATA = _BASE / '抽签排序' / 'data'
DATA_DIR = _DEV_DATA if _DEV_DATA.is_dir() else _BASE / 'data'
DRAW_DIR = DATA_DIR / 'draws'
MUSIC_DIR = DATA_DIR / 'music'
STATE_FILE = DATA_DIR / 'state.json'

for _d in (DATA_DIR, DRAW_DIR, MUSIC_DIR):
    try:
        _d.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass

DEFAULT_STATE = {
    'settings': {'title': '企业宣讲比赛', 'subtitle': '出场顺序抽签',
                 'group_count': 0, 'speed': 1.0,
                 'screen_style': STYLE_DEF_DEFAULT, 'screen_bg': '',
                 'font_family': '', 'text_color': '#ffffff', 'logo_path': ''},
    'participants': [],
    'result': None,
    'music': None,
}


# ---------------------------------------------------------------------------
# 状态持久化
# ---------------------------------------------------------------------------
def load_state():
    state = json.loads(json.dumps(DEFAULT_STATE))
    if STATE_FILE.exists():
        try:
            data = json.loads(STATE_FILE.read_text(encoding='utf-8'))
            state['settings'].update(data.get('settings') or {})
            state['participants'] = data.get('participants') or []
            state['result'] = data.get('result')
            music = data.get('music')
            if music and music.get('filename'):
                if not (MUSIC_DIR / music['filename']).exists():
                    music = None
            state['music'] = music
        except (json.JSONDecodeError, OSError, ValueError):
            pass
    return state


def save_state(state):
    try:
        STATE_FILE.write_text(
            json.dumps(state, ensure_ascii=False, indent=2),
            encoding='utf-8')
    except OSError:
        pass


def _clean_participants(parts):
    out = []
    for p in parts or []:
        name = str(p.get('name') or '').strip()
        if not name:
            continue
        out.append({'name': name,
                    'org': str(p.get('org') or '').strip(),
                    'note': str(p.get('note') or '').strip()})
    return out


# ---------------------------------------------------------------------------
# 名单解析：CSV / Excel / 粘贴文本
# ---------------------------------------------------------------------------
NAME_KEYS = {'姓名', '名字', '选手', '参赛者', '参赛选手', '选手姓名', '人员', 'name'}
ORG_KEYS = {'单位', '部门', '公司', '所属单位', '所属部门', '团队', '队伍', '机构',
            'org', 'organization', 'department', 'team', 'company'}
NOTE_KEYS = {'备注', '说明', 'note', 'remark', '备注信息', '宣讲题目', '题目'}
IGNORE_KEYS = {'序号', '编号', 'no', 'id', 'index', '排名'}


def _normalize(value):
    if value is None:
        return ''
    return str(value).strip()


def _decode_text(raw):
    for enc in ('utf-8-sig', 'utf-8', 'gb18030', 'gbk', 'big5'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ValueError('无法识别文件编码，请另存为 UTF-8 或 GBK 后重试')


def _match_column(header, keys):
    for i, cell in enumerate(header):
        if cell.lower() in keys:
            return i
    return None


def _cell(row, idx):
    if idx is None or idx >= len(row):
        return ''
    return row[idx]


def rows_to_participants(rows):
    """把表格二维数据转成 [{name, org, note}]，自动识别表头。"""
    rows = [[_normalize(c) for c in row] for row in rows]
    rows = [row for row in rows if any(row)]
    if not rows:
        return []

    header = rows[0]
    lowered = [c.lower() for c in header]
    has_header = any(c in NAME_KEYS for c in lowered)

    if has_header:
        name_i = _match_column(header, NAME_KEYS)
        org_i = _match_column(header, ORG_KEYS)
        note_i = _match_column(header, NOTE_KEYS)
        used = {i for i in (name_i, org_i, note_i) if i is not None}
        if name_i is None:
            for i, cell in enumerate(lowered):
                if i not in used and cell not in IGNORE_KEYS:
                    name_i = i
                    break
        body = rows[1:]
    else:
        name_i, org_i, note_i = 0, 1, 2
        body = rows

    if name_i is None:
        return []

    participants = []
    for row in body:
        name = row[name_i] if name_i < len(row) else ''
        if not name:
            continue
        participants.append({'name': name,
                             'org': _cell(row, org_i),
                             'note': _cell(row, note_i)})
    return participants


def parse_upload_file(path):
    """从 xlsx / xlsm / csv / txt 文件解析名单。"""
    suffix = Path(path or '').suffix.lower()
    if suffix in ('.xlsx', '.xlsm'):
        try:
            from openpyxl import load_workbook
        except ImportError:
            raise ValueError('当前环境缺少 openpyxl，无法读取 Excel。\n'
                             '请把名单另存为 CSV / TXT 再导入。')
        wb = load_workbook(path, data_only=True, read_only=True)
        try:
            ws = wb.active
            rows = [[c for c in row] for row in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
        return rows_to_participants(rows)
    if suffix in ('.csv', '.txt'):
        text = _decode_text(Path(path).read_bytes())
        return rows_to_participants(list(csv.reader(io.StringIO(text))))
    raise ValueError('仅支持 .xlsx / .xlsm / .csv / .txt 格式的名单文件')


def parse_pasted_text(text):
    # 制表符（Excel 区域直接粘贴）也按分隔符处理
    text = (text or '').replace('\t', ',')
    return rows_to_participants(list(csv.reader(io.StringIO(text))))


# ---------------------------------------------------------------------------
# 抽签核心（可复现）
# ---------------------------------------------------------------------------
def perform_draw(participants, group_count=0):
    """返回 (seed, items)。items 中 no 即出场签号。"""
    seed = secrets.randbits(64)
    rng = random.Random(seed)
    order = list(range(len(participants)))
    rng.shuffle(order)

    items = []
    for pos, idx in enumerate(order, start=1):
        p = participants[idx]
        items.append({'no': pos,
                     'name': p.get('name', ''),
                     'org': p.get('org', ''),
                     'note': p.get('note', ''),
                     'group': 0})

    # 分组：随机序上按轮转分配，各组人数最多相差 1
    if group_count and group_count > 1:
        for i, item in enumerate(items):
            item['group'] = i % group_count + 1
    return seed, items


def build_record(state, seed, items):
    settings = state['settings']
    return {
        'id': uuid.uuid4().hex[:12],
        'created_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'title': settings.get('title', ''),
        'subtitle': settings.get('subtitle', ''),
        'group_count': settings.get('group_count', 0),
        'seed': seed,
        'algorithm': ALGORITHM,
        'input': _clean_participants(state['participants']),
        'items': items,
    }


def save_record(record):
    path = DRAW_DIR / ('%s.json' % record['id'])
    path.write_text(json.dumps(record, ensure_ascii=False, indent=2),
                    encoding='utf-8')


def list_records(limit=100):
    records = []
    try:
        paths = list(DRAW_DIR.glob('*.json'))
    except OSError:
        return []
    for path in paths:
        try:
            data = json.loads(path.read_text(encoding='utf-8'))
        except (json.JSONDecodeError, OSError, ValueError):
            continue
        records.append({'id': data.get('id'),
                        'created_at': data.get('created_at') or '',
                        'title': data.get('title') or '',
                        'count': len(data.get('items') or []),
                        'seed': data.get('seed')})
    records.sort(key=lambda r: r['created_at'], reverse=True)
    return records[:limit]


def load_record(record_id):
    path = DRAW_DIR / ('%s.json' % Path(str(record_id)).name)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except (json.JSONDecodeError, OSError, ValueError):
        return None


def verify_record(record):
    """用记录的 seed 复现抽签，校验结果与存档是否一致。"""
    if not record:
        return None
    input_list = record.get('input') or []
    rng = random.Random(record.get('seed'))
    order = list(range(len(input_list)))
    rng.shuffle(order)
    reproduced = [input_list[i].get('name', '') for i in order]
    actual = [item.get('name', '') for item in record.get('items', [])]
    return reproduced == actual


# ---------------------------------------------------------------------------
# 导出
# ---------------------------------------------------------------------------
def build_xlsx_workbook(record):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = '出场顺序'

    group_count = record.get('group_count') or 0
    headers = ['出场签号', '姓名', '单位 / 部门']
    if group_count and group_count > 1:
        headers.append('组别')
    headers.append('备注')

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    c = ws.cell(row=1, column=1, value=record.get('title', '') or '抽签结果')
    c.font = Font(size=16, bold=True)
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    meta = ('%s　|　抽签时间：%s　|　随机种子：%s　|　记录编号：%s'
            % (record.get('subtitle', ''), record.get('created_at', ''),
               record.get('seed', ''), record.get('id', '')))
    c = ws.cell(row=2, column=1, value=meta.strip('　| '))
    c.font = Font(size=10, color='666666')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 22

    head_fill = PatternFill('solid', fgColor='1F3864')
    thin = Side(style='thin', color='BFBFBF')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, name in enumerate(headers, start=1):
        c = ws.cell(row=3, column=col, value=name)
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = head_fill
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = border
    ws.row_dimensions[3].height = 24

    for r, item in enumerate(record.get('items', []), start=4):
        values = [item.get('no'), item.get('name'), item.get('org')]
        if group_count and group_count > 1:
            values.append(item.get('group'))
        values.append(item.get('note'))
        for col, value in enumerate(values, start=1):
            c = ws.cell(row=r, column=col, value=value)
            c.border = border
            c.alignment = Alignment(
                horizontal='center' if col != 3 else 'left', vertical='center')
        ws.row_dimensions[r].height = 22

    widths = [12, 18, 28]
    if group_count and group_count > 1:
        widths += [10, 28]
    else:
        widths += [28]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return wb


def export_xlsx(record, path):
    build_xlsx_workbook(record).save(path)


def build_csv_text(record):
    group_count = record.get('group_count') or 0
    header = ['出场签号', '姓名', '单位/部门']
    if group_count and group_count > 1:
        header.append('组别')
    header.append('备注')

    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow([record.get('title', '')])
    writer.writerow(['%s 抽签时间:%s 随机种子:%s 记录编号:%s'
                     % (record.get('subtitle', ''), record.get('created_at', ''),
                        record.get('seed', ''), record.get('id', ''))])
    writer.writerow(header)
    for item in record.get('items', []):
        row = [item.get('no'), item.get('name'), item.get('org')]
        if group_count and group_count > 1:
            row.append(item.get('group'))
        row.append(item.get('note'))
        writer.writerow(row)
    return '\ufeff' + buf.getvalue()


# ---------------------------------------------------------------------------
# 声音：winsound 蜂鸣（零依赖）+ MCI 背景音乐
# ---------------------------------------------------------------------------
def _beep_thread(freqs):
    def run():
        try:
            import winsound
            for freq, dur in freqs:
                winsound.Beep(freq, dur)
        except Exception:
            pass
    threading.Thread(target=run, daemon=True).start()


def blip():
    _beep_thread([(880, 110), (1320, 70)])


def fanfare():
    _beep_thread([(523, 220), (659, 220), (784, 220), (1046, 420)])


class MciPlayer(object):
    """背景音乐播放器：优先用 WMPlayer.OCX（流式缓冲，不卡顿），
    失败时回退到 winmm MCI。支持 mp3/wav 等常见格式循环播放。"""

    def __init__(self, alias):
        self._alias = alias
        self._playing = False
        self._wmp = None          # WMPlayer.OCX 实例
        self._repeat = True

    # ---------- WMP 实现（首选） ----------
    def _ensure_wmp(self):
        if self._wmp is not None:
            return True
        try:
            import win32com.client
            self._wmp = win32com.client.Dispatch('WMPlayer.OCX.7')
            return True
        except Exception:
            self._wmp = None
            return False

    def _wmp_play(self, path, repeat):
        try:
            # 清掉上一首，避免 setMode/URL 顺序在缓存曲目时偶发不生效
            self._wmp.URL = ''
            self._wmp.settings.setMode('loop', bool(repeat))
            self._wmp.settings.volume = 100
            self._wmp.URL = path
            # WMP 打开 URL 后会自行缓冲并开始播放，无需轮询阻塞 UI
            return True
        except Exception:
            return False

    def _wmp_stop(self):
        try:
            self._wmp.controls.stop()
            self._wmp.URL = ''
        except Exception:
            pass

    # ---------- MCI 兜底实现 ----------
    def _mci_send(self, cmd):
        try:
            import ctypes
            winmm = ctypes.WinDLL('winmm')
            send = winmm.mciSendStringW
            buf = ctypes.create_unicode_buffer(256)
            return send(cmd, buf, 254, 0)
        except Exception:
            return 1

    def _mci_play(self, path, repeat):
        self._mci_send('close %s' % self._alias)
        if self._mci_send('open "%s" alias %s'
                          % (path.replace('"', ''), self._alias)) != 0:
            return False
        cmd = 'play %s%s' % (self._alias, ' repeat' if repeat else '')
        if self._mci_send(cmd) != 0:
            self._mci_send('close %s' % self._alias)
            return False
        return True

    def _mci_stop(self):
        self._mci_send('close %s' % self._alias)

    # ---------- 对外接口 ----------
    def play(self, path, repeat=True):
        self.stop()
        path = str(Path(path).resolve())
        if not os.path.exists(path):
            return False
        self._repeat = bool(repeat)
        if self._ensure_wmp() and self._wmp_play(path, repeat):
            self._playing = True
            return True
        # WMP 不可用，回退 MCI
        if self._mci_play(path, repeat):
            self._playing = True
            return True
        return False

    def stop(self):
        if self._wmp is not None:
            self._wmp_stop()
        self._mci_stop()
        self._playing = False

    @property
    def playing(self):
        if not self._playing:
            return False
        # WMP 实际已播完（非循环）时，playState != 3
        if self._wmp is not None:
            try:
                return self._wmp.playState == 3
            except Exception:
                pass
        return self._playing


# ---------------------------------------------------------------------------
# 多显示器枚举（大屏投屏用；用 WINFUNCTYPE 私有副本，不污染共享函数对象）
# ---------------------------------------------------------------------------
def _enum_monitors():
    """枚举全部显示器：[{'hmon', 'rect': (x,y,w,h), 'primary'}]。"""
    try:
        import ctypes
        from ctypes import wintypes

        class _MONITORINFO(ctypes.Structure):
            _fields_ = [('cbSize', wintypes.DWORD),
                        ('rcMonitor', wintypes.RECT),
                        ('rcWork', wintypes.RECT),
                        ('dwFlags', wintypes.DWORD)]

        out = []
        user32 = ctypes.WinDLL('user32')
        user32.GetMonitorInfoW.argtypes = [wintypes.HMONITOR,
                                           ctypes.c_void_p]
        user32.GetMonitorInfoW.restype = wintypes.BOOL
        proto = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HMONITOR,
                                   wintypes.HDC, wintypes.LPRECT,
                                   wintypes.LPARAM)

        def _cb(hmon, _hdc, _lprect, _lparam):
            info = _MONITORINFO()
            info.cbSize = ctypes.sizeof(_MONITORINFO)
            if user32.GetMonitorInfoW(hmon, ctypes.byref(info)):
                r = info.rcMonitor
                out.append({'hmon': int(hmon) if hmon else 0,
                            'rect': (r.left, r.top,
                                     r.right - r.left, r.bottom - r.top),
                            'primary': bool(info.dwFlags & 1)})
            return True

        user32.EnumDisplayMonitors(None, None, proto(_cb), 0)
        return out
    except Exception:
        return []


def list_monitors():
    return [m['rect'] for m in _enum_monitors()]


def _window_monitor(master):
    """返回管理窗所在显示器的 HMONITOR 句柄。

    直接用窗口句柄向系统查询，不做坐标包含判断，避免多屏不同 DPI 时
    Tk 逻辑坐标与显示器物理坐标不一致而误判。
    """
    try:
        import ctypes
        user32 = ctypes.WinDLL('user32')
        user32.MonitorFromWindow.restype = ctypes.c_void_p
        hwnd = master.winfo_id()
        hmon = user32.MonitorFromWindow(hwnd, 2)  # MONITOR_DEFAULTTONEAREST
        if not hmon:  # Tk 子窗口查不到时退到外层包装窗口
            parent = user32.GetParent(hwnd)
            hmon = user32.MonitorFromWindow(parent or hwnd, 2)
        return int(hmon) if hmon else None
    except Exception:
        return None


def _pick_monitor(master):
    """选择投屏显示器：默认投到管理窗所在屏幕之外的那块（扩展屏/投影仪）。

    管理窗通常在主屏，因此多屏时即默认落到扩展屏；只有一块屏时留在
    原屏。句柄判定失败时退化为“优先非主屏”。
    """
    mons = _enum_monitors()
    if not mons:
        w = master.winfo_screenwidth()
        h = master.winfo_screenheight()
        return 0, 0, w, h
    cur_h = _window_monitor(master)
    if cur_h is not None:
        others = [m for m in mons if m['hmon'] != cur_h]
    else:
        others = [m for m in mons if not m['primary']]
    if others:
        # 多块扩展屏时非主屏优先（顺序稳定，主屏始终排最后）
        others.sort(key=lambda m: m['primary'])
        return others[0]['rect']
    cur = next((m for m in mons if m['hmon'] == cur_h), None)
    return (cur or mons[0])['rect']


def _safe_no_topmost(w):
    """安全地解除 topmost（_keep_front 的延时回调）。"""
    try:
        if w.winfo_exists():
            w.attributes('-topmost', False)
    except Exception:
        pass


def _is_hex(v):
    """粗略校验 #rgb / #rrggbb 颜色。"""
    if not isinstance(v, str) or not v:
        return False
    if v[0] != '#' or len(v) not in (4, 7):
        return False
    try:
        int(v[1:], 16)
        return True
    except ValueError:
        return False


def ui_theme_families():
    """返回常见中英文字体候选，只留系统里真实存在的。"""
    wanted = ('微软雅黑', 'Microsoft YaHei UI', 'Microsoft YaHei', '黑体',
              'SimHei', '宋体', 'SimSun', '楷体', 'KaiTi', '隶书',
              'YouYuan', '幼圆', '等线', 'DengXian', '方正舒体',
              'FZShuTi', 'Arial', 'Tahoma', 'Segoe UI', 'Times New Roman',
              'Arial Black', 'Impact', 'Cambria', 'Georgia')
    try:
        have = set(tkfont.families())
    except Exception:
        have = set()
    out = [n for n in wanted if n in have]
    return out or ['']


# measure/fit 结果缓存，避免大屏渲染时对同一 (文本, 宽度) 反复测量，
# tkfont.measure 是 Tcl 往返调用，线性逐号缩小会成为打开大屏的卡顿热点
_FIT_CACHE = {}


def _fit_font(master, text, max_w, start_size, min_size, weight='normal'):
    """在 [min, start] 里用二分找能塞进 max_w 的最大字号，结果缓存。"""
    key = (text, max_w, start_size, min_size, weight)
    if key in _FIT_CACHE:
        size = _FIT_CACHE[key]
    else:
        lo, hi = min_size, start_size
        best = None
        while lo <= hi:
            mid = (lo + hi) // 2
            if tkfont.Font(
                    master=master, font=font(mid, weight)).measure(text) <= max_w:
                best, lo = mid, mid + 1
            else:
                hi = mid - 1
        size = best if best is not None else min_size
        if len(_FIT_CACHE) > 2000:
            _FIT_CACHE.clear()
        _FIT_CACHE[key] = size
    return tkfont.Font(master=master, font=font(size, weight))


# ---------------------------------------------------------------------------
# 大屏投屏窗口
# ---------------------------------------------------------------------------
class LotteryScreen(object):
    """大屏窗口：卡片网格 + 轮播 → 逐个定格 → 撒花 + 音效。"""

    BADGE = {'idle': ('待抽签', '#8ea0c8'),
             'rolling': ('抽签中', '#7cb9ff'),
             'locking': ('揭晓中', '#ffc46b'),
             'done': ('已完成', '#7cb342')}

    def __init__(self, master, admin):
        self.admin = admin
        self.win = tk.Toplevel(master)
        # Windows 上 Toplevel 创建即可见（默认小窗落在左上角），若先显示
        # 再去边框+跨屏定位，会看到“内容堆在左上角再展开”。先隐藏，
        # 定位/首帧绘制完成后再一次性显示。
        self.win.withdraw()
        self.win.title('抽签大屏')
        self.win.configure(bg=SC_BG)
        self.mode = 'idle'
        self.participants = []
        self.items = []
        self.record = None
        self.reveal = 0
        self.auto = True
        self.sound_on = True
        self.music_on = True
        self._rolling_names = []
        self._roll_tick = 0
        self._cycle_job = None
        self._step_job = None
        self._keep_job = None
        # 抽签速度参数（start_draw 时按后台设置倍率重算）
        self._roll_ms = ROLL_MS
        self._cycle_ms = CYCLE_MS
        self._auto_total_ms = AUTO_TOTAL_MS
        self._confetti = []          # 礼花粒子
        self._rings = []             # 金色光环
        self._confetti_job = None    # 统一特效帧循环
        self._fx_wave_jobs = []      # 礼花分波 / 光环错峰的延迟任务
        self._toast_job = None
        self._toast_msg = None        # 当前存活的 toast 文案（_render 后自动恢复）
        self._toast_img = None
        self._full = True
        self._mci = MciPlayer('lottery_bgm')
        self._bg_photo = None      # 自定义背景位图（按尺寸缓存）
        self._bg_key = None
        self._logo_photo = None    # 上传 Logo 位图（按尺寸缓存）
        self._logo_key = None
        self._fam = ''             # 大屏自定义字体族
        self._main_col = SC_TEXT   # 大屏主要文字颜色
        self._closed = False
        # idle 态「待抽签」虚化位图：按尺寸缓存，当前渲染引用另外持有
        self._blur_cache = {}
        self._blur_refs = []
        # 轮播文字约束（Canvas 文字不会被卡片裁剪，需自行按宽度截断）
        self._rolling_font = None
        self._cycle_w = 0
        self._cycle_fit_cache = {}

        self.canvas = tk.Canvas(self.win, bg=SC_BG, highlightthickness=0, bd=0)
        self.canvas.pack(fill='both', expand=True)
        # 不用"点击画布即启动/推进"，防止大屏现场误触；触发统一走空格键
        self.canvas.bind('<Configure>', self._on_configure)
        self._last_size = None
        self._size_job = None

        self.win.bind('<space>', self._on_space)
        self.win.bind('<Escape>', self._on_escape)
        self.canvas.bind('<Escape>', self._on_escape)
        for ch, handler in (('r', self._on_r), ('f', self._on_f),
                            ('m', self._on_m), ('b', self._on_b),
                            ('e', self._on_e)):
            self.win.bind('<%s>' % ch, handler)
            self.win.bind('<Key-%s>' % ch.upper(), handler)

        self.win.protocol('WM_DELETE_WINDOW', self.close)
        self._mx, self._my, self._mw, self._mh = _pick_monitor(master)
        self._apply_full()
        # 隐藏状态下先让布局生效，画布在首帧绘制前就拿到全屏真实尺寸，
        # 避免先按小窗画一版、显示后再重排。
        self.win.update_idletasks()

        self.participants = _clean_participants(admin.state.get('participants'))
        result = admin.state.get('result')
        if result and result.get('items'):
            self._restore_result(result)
        else:
            self._render()

        self.win.deiconify()
        ui_theme.present_window(self.win)
        try:
            self.win.focus_force()
        except Exception:
            pass
        self.win.after(260, lambda: self._safe(ui_theme.present_window))
        self._keep_alive()

    # ---------- 窗口管理 ----------
    def _safe(self, fn, *args):
        try:
            if self.win.winfo_exists():
                fn(self.win, *args)
        except Exception:
            pass

    def _apply_full(self):
        try:
            if self._full:
                self.win.overrideredirect(True)
                self.win.geometry('%dx%d+%d+%d' % (self._mw, self._mh,
                                                   self._mx, self._my))
            else:
                self.win.overrideredirect(False)
                self.win.state('zoomed')
        except Exception:
            pass
        ui_theme.present_window(self.win)

    def _keep_alive(self):
        """无边框窗口投到扩展屏时系统可能不发重画消息，定期强制重画。"""
        try:
            if not self.win.winfo_exists() or self._closed:
                return
            # 动画进行中不强制重画，避免强重画与轮播/撒花叠加把主线程占满成假死；
            # 动画结束后再恢复 1.5s 一次的正常强重画，保证扩展屏不冻结。
            if self._cycle_job or self._step_job or self._confetti_job:
                delay = 3000
            else:
                ui_theme.force_repaint(self.win)
                delay = 1500
            self._keep_job = self.win.after(delay, self._keep_alive)
        except Exception:
            pass

    def close(self):
        if self._closed:
            return
        self._closed = True
        # 大屏字体的全局覆盖只在展示期间生效：关闭即复位，避免污染
        # 之后重建的首页/后台等其它界面字体
        ui_theme.set_font_override('')
        self._stop_jobs()
        self._mci.stop()
        try:
            if self.win.winfo_exists():
                self.win.destroy()
        except Exception:
            pass

    def _stop_jobs(self):
        for attr in ('_cycle_job', '_step_job', '_keep_job',
                     '_confetti_job', '_toast_job', '_size_job'):
            job = getattr(self, attr, None)
            if job:
                try:
                    self.win.after_cancel(job)
                except Exception:
                    pass
                setattr(self, attr, None)
        for job in getattr(self, '_fx_wave_jobs', []):
            try:
                self.win.after_cancel(job)
            except Exception:
                pass
        self._fx_wave_jobs = []

    def is_open(self):
        try:
            return bool(self.win.winfo_exists()) and not self._closed
        except Exception:
            return False

    # ---------- 键盘 ----------
    def _on_space(self, _e=None):
        self.advance()
        return 'break'

    def _on_escape(self, _e=None):
        # 退出大屏：只关投屏窗口，管理窗口和抽签数据都保留
        self.close()
        return 'break'

    def _on_r(self, _e=None):
        # 重置只能由管理后台执行，避免大屏现场误触把结果清掉
        self.toast('请在管理后台点「重置结果」清除本次抽签')
        return 'break'

    def _on_f(self, _e=None):
        self._full = not self._full
        self._apply_full()
        self._render()
        return 'break'

    def _on_m(self, _e=None):
        self.sound_on = not self.sound_on
        if self.sound_on:
            blip()
        self._render()
        return 'break'

    def _on_b(self, _e=None):
        if not self.admin.get_music_path():
            self.toast('还没有设置背景音乐')
            return 'break'
        self.music_on = not self.music_on
        if self.music_on:
            path = self.admin.get_music_path()
            if path and self.mode in ('rolling', 'locking', 'done'):
                self._mci.play(path)
        else:
            self._mci.stop()
        self._render()
        return 'break'

    def _on_e(self, _e=None):
        self.admin.export_result(fmt='xlsx')
        return 'break'

    # ---------- 绘制 ----------
    def _restore_result(self, record):
        self.record = record
        self.items = record.get('items') or []
        self.reveal = len(self.items)
        self.mode = 'done'
        self._render()

    def _on_configure(self, _e=None):
        """窗口尺寸变化：合并连续触发 + 尺寸真变才重绘，避免扩展屏抖动触发整屏重建卡顿。"""
        if self._closed:
            return
        if self._size_job:
            try:
                self.win.after_cancel(self._size_job)
            except Exception:
                pass
        self._size_job = self.win.after(120, self._do_render_if_resized)

    def _do_render_if_resized(self):
        self._size_job = None
        try:
            if self._closed or not self.win.winfo_exists():
                return
            w = self.canvas.winfo_width()
            h = self.canvas.winfo_height()
        except Exception:
            return
        if (w, h) == self._last_size:
            return
        self._last_size = (w, h)
        self._render()

    def _refresh_chrome(self):
        """揭晓切档/完成时只更新状态徽章与底部提示，不整屏重建，避免逐张定格假死。"""
        try:
            if self._closed or not self.win.winfo_exists():
                return
            c = self.canvas
            w = max(c.winfo_width(), 100)
            h = max(c.winfo_height(), 100)
            c.delete('badge')
            c.delete('foot')
            _label, _color = self.BADGE.get(self.mode, self.BADGE['idle'])
            _bw, _bh = 118, 34
            round_rect(c, w - _bw - 28, 24, w - 28, 24 + _bh, 17,
                       fill=SC_BG2, outline=_color, width=1, tags=('badge',))
            c.create_text(w - 28 - _bw // 2, 24 + _bh // 2 + 1,
                          text=_label, font=font(13, 'bold'),
                          fill=_color, tags=('badge',))
            _foot = self._foot_text()
            _foot_fill = '#46557d' if self.mode == 'done' else SC_MUTED
            c.create_text(w // 2, h - 26, text=_foot,
                          font=font(12), fill=_foot_fill, tags=('foot',))
        except tk.TclError:
            pass

    def _apply_style_prefs(self):
        """读设置里的字体族/文字颜色，应用到当前大屏渲染。"""
        st = self.admin.state.get('settings') or {}
        self._fam = st.get('font_family') or ''
        ui_theme.set_font_override(self._fam)
        col = st.get('text_color')
        self._main_col = col if _is_hex(col) else SC_TEXT

    def _bg_photo_for(self, w, h):
        """返回自适应画布尺寸的自定义背景位图（居中 cover 填满 + 压暗）。

        按 (路径,尺寸) 缓存与刷新：画布尺寸未变时复用，改名/改图时重建。
        """
        path = (self.admin.state.get('settings') or {}).get('screen_bg')
        if not path or not os.path.exists(path):
            self._bg_photo = None
            self._bg_key = None
            return None
        key = (path, w, h)
        if self._bg_photo is not None and self._bg_key == key:
            return self._bg_photo
        try:
            from PIL import Image, ImageTk
            img = Image.open(path).convert('RGB')
            iw, ih = img.size
            scale = max(w / iw, h / ih)
            nw, nh = max(1, int(iw * scale)), max(1, int(ih * scale))
            # BILINEAR 远快于 LANCZOS，背景本就压暗 50%，画质损失无感知，
            # 避免大图/超大屏首帧卡顿
            img = img.resize((nw, nh), Image.BILINEAR)
            l = (nw - w) // 2
            t = (nh - h) // 2
            img = img.crop((l, t, l + w, t + h))
            # 整体压暗，保证标题/卡片/文字清晰可读
            shade = Image.new('RGB', img.size, (10, 16, 34))
            img = Image.blend(img, shade, 0.5)
            self._bg_photo = ImageTk.PhotoImage(img, master=self.canvas)
            self._bg_key = key
            return self._bg_photo
        except Exception:
            self._bg_photo = None
            self._bg_key = None
            return None

    def _draw_logo(self, c, w, h):
        """左上角绘制上传的 Logo（保持宽高比，按高度收纳）。"""
        path = (self.admin.state.get('settings') or {}).get('logo_path')
        if not path or not os.path.exists(path):
            return
        key = (path, int(w), int(h))
        if self._logo_photo is not None and self._logo_key == key:
            photo = self._logo_photo
        else:
            try:
                from PIL import Image, ImageTk
                img = Image.open(path).convert('RGBA')
                lw, lh = img.size
                if not lw or not lh:
                    return
                th = max(24, int(h * 0.08))
                scale = th / float(lh)
                nw, nh = int(lw * scale), th
                if nw > int(w * 0.28):          # 太宽则改用宽度约束
                    scale = (w * 0.28) / float(lw)
                    nw, nh = int(lw * scale), int(lh * scale)
                img = img.resize((max(1, nw), max(1, nh)), Image.LANCZOS)
                self._logo_photo = ImageTk.PhotoImage(img, master=c)
                self._logo_key = key
                photo = self._logo_photo
            except Exception:
                return
        c.create_image(20, 20, image=photo, anchor='nw')

    def _render(self):
        try:
            if self._closed or not self.win.winfo_exists():
                return
            c = self.canvas
            w = max(c.winfo_width(), 100)
            h = max(c.winfo_height(), 100)
            c.delete('all')
            # 旧画布项已删，只保留当前这一屏实际显示的虚化位图引用
            self._blur_refs = []

            settings = self.admin.state.get('settings') or {}
            self._apply_style_prefs()
            # 自定义大屏背景：置于最底层（先画），覆盖全部画布
            bg_img = self._bg_photo_for(w, h)
            if bg_img is not None:
                c.create_image(w // 2, h // 2, image=bg_img)
            # 上传的 Logo：左上角，保持宽高比
            self._draw_logo(c, w, h)
            # 标题 / 副标题 / 元信息：按字体真实行高纵向堆叠，大分辨率下
            # 也不互相压线；标题字号封顶，避免大屏上标题过大
            title_f = tkfont.Font(
                master=c,
                font=font(max(20, min(56, int(h * 0.042))), 'bold'))
            sub_f = tkfont.Font(
                master=c, font=font(max(12, min(24, int(h * 0.022)))))
            meta_f = tkfont.Font(
                master=c, font=font(max(10, min(18, int(h * 0.018)))))
            y = int(h * 0.03) + title_f.metrics('linespace') // 2
            c.create_text(w // 2, y,
                          text=settings.get('title') or '出场顺序抽签',
                          font=title_f, fill=self._main_col)
            y += title_f.metrics('linespace') // 2 + 6
            if settings.get('subtitle'):
                y += sub_f.metrics('linespace') // 2
                c.create_text(w // 2, y, text=settings['subtitle'],
                              font=sub_f, fill=SC_MUTED)
                y += sub_f.metrics('linespace') // 2 + 8

            # 状态徽章（右上角）
            label, color = self.BADGE.get(self.mode, self.BADGE['idle'])
            bw, bh = 118, 34
            round_rect(c, w - bw - 28, 24, w - 28, 24 + bh, 17,
                       fill=SC_BG2, outline=color, width=1, tags=('badge',))
            c.create_text(w - 28 - bw // 2, 24 + bh // 2 + 1, text=label,
                          font=font(13, 'bold'), fill=color, tags=('badge',))

            # 元信息行（接在副标题下方）
            meta = self._meta_text()
            if meta:
                y += meta_f.metrics('linespace') // 2
                c.create_text(w // 2, y, text=meta,
                              font=meta_f, fill=SC_MUTED)

            n = len(self.items) or len(self.participants)
            if n:
                # 网格从标题/副标题/元信息实际底部再加留白开始，
                # 避免首行卡片压到上方小标题
                if meta:
                    y += meta_f.metrics('linespace') // 2
                grid_top = y + max(10, min(24, int(h * 0.02)))
                self._render_cards(c, w, h, n, top=grid_top)

            # 底部提示
            foot = self._foot_text()
            foot_fill = '#46557d' if self.mode == 'done' else SC_MUTED
            c.create_text(w // 2, h - 26, text=foot,
                          font=font(12), fill=foot_fill, tags=('foot',))

            # idle 态：居中操作提示浮层（对齐旧版网页大屏的 CTA 卡片）
            if self.mode == 'idle':
                self._draw_cta(c, w, h, n > 0)

            # 特效在运行中（如 F 切换全屏触发重绘）：按当前状态重建画布项
            if self._confetti or self._rings:
                self._redraw_fx(c, w, h)

            # 若有存活的 toast，整屏重绘后恢复（不重置倒计时）
            if self._toast_msg:
                self._draw_toast(self._toast_msg)
        except tk.TclError:
            pass

    def _foot_text(self):
        if self.mode == 'done':
            return '抽签已完成 · 重抽请回管理后台点「重置结果」 · ESC 退出大屏'
        return ('空格 开始/推进 · F 全屏 · M 音效%s · E 导出 · ESC 退出'
                % (' · B 音乐' if self.admin.get_music_path() else ''))

    def _draw_cta(self, c, w, h, has_people):
        """idle 态居中的操作提示浮层（对齐旧版网页大屏的 CTA 卡片）。

        金色主提示 + 一行带按键小框（kbd）的快捷键说明；抽签开始后
        _render 不再绘制本浮层，等同于旧版的淡出隐藏。
        """
        main_f = tkfont.Font(
            master=c, font=font(max(18, min(34, int(h * 0.028))), 'bold'))
        sub_f = tkfont.Font(
            master=c, font=font(max(12, min(17, int(h * 0.015)))))
        key_f = tkfont.Font(
            master=c, font=font(max(12, min(17, int(h * 0.015))), 'bold'))

        if has_people:
            main_text = '准备就绪 · 抽取出场顺序'
            segments = [
                ('t', '按 '), ('k', '空格'), ('t', ' 开始　·　'),
                ('k', '空格'), ('t', ' 逐个揭晓　·　'),
                ('k', 'B'), ('t', ' 音乐　·　'),
                ('k', 'E'), ('t', ' 导出　·　'),
                ('k', 'ESC'), ('t', ' 退出'),
            ]
        else:
            main_text = '尚未导入名单 · 请先在管理窗口添加参赛者'
            segments = [('k', 'ESC'), ('t', ' 退出大屏')]

        # 先逐段排版，算出副行总宽
        kbd_px, kbd_py = 12, 5
        kbd_h = key_f.metrics('linespace') + kbd_py * 2
        laid = []
        x = 0
        for kind, text in segments:
            tw = key_f.measure(text) if kind == 'k' else sub_f.measure(text)
            if kind == 'k':
                sw, sh = tw + kbd_px * 2, kbd_h
            else:
                sw, sh = tw, sub_f.metrics('linespace')
            laid.append((kind, text, x, sw, sh))
            x += sw + 2
        sub_w = x - 2

        pad_x = max(28, int(h * 0.028))
        pad_y = max(16, int(h * 0.022))
        gap_mid = max(12, int(h * 0.012))
        main_lh = main_f.metrics('linespace')
        box_w = min(max(main_f.measure(main_text), sub_w) + pad_x * 2,
                    int(w * 0.9))
        box_h = pad_y * 2 + main_lh + gap_mid + kbd_h
        cx = w // 2
        cy = int(h * 0.56)
        x1, y1 = cx - box_w // 2, cy - box_h // 2
        x2, y2 = cx + box_w // 2, cy + box_h // 2
        round_rect(c, x1, y1, x2, y2, 20,
                   fill='#0d1830', outline=SC_LINE, width=1)

        ty = y1 + pad_y + main_lh // 2
        c.create_text(cx, ty, text=main_text, font=main_f, fill=SC_FLASH)

        sx = cx - sub_w // 2
        ky = ty + main_lh // 2 + gap_mid + kbd_h // 2
        for kind, text, ox, sw, _sh in laid:
            if kind == 'k':
                rx1 = sx + ox
                ry1 = ky - kbd_h // 2
                round_rect(c, rx1, ry1, rx1 + sw, ry1 + kbd_h, 7,
                           fill='#1b294a', outline=SC_LINE, width=1)
                c.create_text(rx1 + sw // 2, ky, text=text,
                              font=key_f, fill=SC_TEXT)
            else:
                c.create_text(sx + ox + sw // 2, ky, text=text,
                              font=sub_f, fill=SC_MUTED)

    def _meta_text(self):
        parts = []
        if self.record:
            count = len(self.items) or len(self.participants)
            parts.append('参赛 %d 人' % count)
            gc = self.record.get('group_count') or 0
            if gc and gc > 1:
                parts.append('分 %d 组' % gc)
            parts.append('记录 %s' % self.record.get('id', ''))
            parts.append('种子 %s' % self.record.get('seed', ''))
            parts.append(self.record.get('created_at', ''))
        else:
            count = len(self.participants)
            if count:
                parts.append('参赛 %d 人' % count)
        return '　·　'.join(parts)

    def _blur_photo(self, text, tk_fnt):
        """把文字渲染成「半透明 + 高斯模糊」位图（对齐旧版网页大屏 idle
        卡片的 opacity:0.3 + filter:blur(3px)，传达"尚未决定"）。

        按像素尺寸缓存；Pillow 不可用时返回 None，由调用方退化为实字。
        """
        try:
            line_px = int(tk_fnt.metrics('linespace'))
            text_px = max(1, int(tk_fnt.measure(text)))
            dpi = int(self.canvas.winfo_fpixels('1i') + 0.5)
        except Exception:
            return None
        key = (text, line_px, text_px, dpi)
        photo = self._blur_cache.get(key)
        if photo is not None:
            return photo
        try:
            from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageTk
        except Exception:
            return None
        try:
            pil_font = None
            for name in ('msyh.ttc', 'msyhui.ttf', 'msyhbd.ttc', 'simhei.ttf'):
                try:
                    pil_font = ImageFont.truetype(name, line_px)
                    break
                except Exception:
                    continue
            if pil_font is None:
                return None
            pad = max(10, line_px // 2)
            img = Image.new('RGBA',
                            (text_px + pad * 2, line_px + pad * 2),
                            (0, 0, 0, 0))
            d = ImageDraw.Draw(img)
            r16, g16, b16 = self.win.winfo_rgb(SC_MUTED)
            color = (r16 >> 8, g16 >> 8, b16 >> 8, 77)  # alpha≈0.30
            try:
                asc, desc = pil_font.getmetrics()
                ty = pad + (line_px - asc - desc) // 2
            except Exception:
                ty = pad
            d.text((pad, ty), text, font=pil_font, fill=color)
            # blur(3px) 是 CSS 像素概念，随实际字号等比放大
            radius = max(2.5, line_px * 0.055)
            img = img.filter(ImageFilter.GaussianBlur(radius=radius))
            photo = ImageTk.PhotoImage(img, master=self.win)
        except Exception:
            return None
        self._blur_cache[key] = photo
        return photo

    def _render_cards(self, c, w, h, n, top=None):
        """全量布局绘制（开窗 / 缩放 / 模式切换时调用一次）。

        动画过程中的轮播和定格只做 itemconfig 增量更新（见
        _cycle_tick / _apply_locked），绝不在每帧重排重绘。
        """
        mgx = max(40, int(w * 0.045))
        top = top if top is not None else int(h * 0.17)
        bottom = h - 92
        avail_w = w - 2 * mgx
        avail_h = max(40, bottom - top)

        # 列数按“目标卡高/卡宽比”随真实屏幕宽高比自适应（而非只看人数）：
        #   cw≈avail_w/cols, ch≈avail_h/rows, rows≈n/cols
        #   ch/cw ≈ avail_h·cols² / (avail_w·n) ≈ RATIO  → 解出 cols。
        # 宽屏(21:9)自动加列摊平、竖屏/窄屏自动减列拉高卡，贴合屏幕。
        RATIO = 1.35
        cols = int(round((RATIO * avail_w * n / max(1, avail_h)) ** 0.5))
        cols = max(1, min(n, cols))
        cols = max(1, min(10, cols))

        # 宽度装不下列数就逐列收拢，保证卡宽至少可读
        while cols > 1 and (avail_w - 14 * (cols - 1)) // cols < 68:
            cols -= 1
        rows = (n + cols - 1) // cols

        # 间隙随画布缩放，屏太小/人数太多时继续收缩，
        # 保证网格始终整块落在画布内（任何分辨率下都不裁切）
        gap = max(2, min(16, int(h * 0.016)))
        if gap * (rows - 1) + rows > avail_h:
            gap = max(2, (avail_h - rows) // max(1, rows - 1))
        cw = max(1, (avail_w - gap * (cols - 1)) // cols)
        ch = max(1, (avail_h - gap * (rows - 1)) // rows)

        grid_w = cols * cw + (cols - 1) * gap
        grid_h = rows * ch + (rows - 1) * gap
        ox = (w - grid_w) // 2
        oy = top + max(0, (bottom - top - grid_h) // 3)

        # 轮播用统一字体（每帧只改文字，不重新 fit）
        # 对齐旧版网页 .card-name 的 18cqh（卡高 18%），开始抽签时
        # 名字不再过大；idle 虚化「待抽签」共用同一字号
        rolling_font = tkfont.Font(
            master=c, font=font(max(13, min(32, int(ch * 0.18)))))
        # 记录给轮播循环：名字必须收进卡片内（Canvas 文字无裁剪，
        # 不截断就会铺到卡片外，快速换帧时表现为框外乱跳）
        self._rolling_font = rolling_font
        self._cycle_w = cw - 24
        self._cycle_fit_cache = {}

        # idle 态未定格用虚化「待抽签」占位；rolling 态文本项可见、由
        # cycle() 播放转动字符动效（不出现真实姓名）
        use_placeholder = self.mode == 'idle'
        placeholder_photo = self._blur_photo('待抽签', rolling_font) \
            if use_placeholder else None

        self._cards = []
        for i in range(n):
            col, row = i % cols, i // cols
            x1 = ox + col * (cw + gap)
            y1 = oy + row * (ch + gap)
            x2, y2 = x1 + cw, y1 + ch

            item = self.items[i] if i < len(self.items) else None
            locked = bool(item) and i < self.reveal

            card = {'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2,
                    'ch': ch, 'cw': cw, 'locked': locked, 'font': None}
            card['rect'] = round_rect(
                c, x1, y1, x2, y2, 14,
                fill=SC_CARD_LOCK if locked else SC_CARD,
                outline=SC_LINE_LOCK if locked else SC_LINE,
                width=2 if locked else 1)

            no_size = max(10, min(24, int(ch * 0.14)))
            card['no'] = c.create_text(
                (x1 + x2) // 2, y1 + int(ch * 0.14),
                text='第 %d 位' % (i + 1),
                font=font(no_size, 'bold'),
                fill=ACCENT if locked else SC_MUTED)

            if locked:
                card['name'] = c.create_text(
                    (x1 + x2) // 2, y1 + int(ch * 0.52),
                    text=item.get('name', ''), font=rolling_font,
                    fill=SC_TEXT)
                card['org'] = c.create_text(
                    (x1 + x2) // 2, y1 + int(ch * 0.78), text='',
                    font=font(max(10, min(18, int(ch * 0.11)))),
                    fill=SC_MUTED)
                self._cards.append(card)
                self._apply_locked(i, flash=False)
            else:
                # 未定格：idle 用虚化「待抽签」占位；rolling 按所选样式播动效；
                # 定格时才写入最终名字并清理占位元素
                rolling = self.mode == 'rolling'
                style = (self.admin.state.get('settings') or {}).get(
                    'screen_style') or STYLE_DEF_DEFAULT
                use_digits = rolling and style == STYLE_DIGITS
                card['name'] = c.create_text(
                    (x1 + x2) // 2, y1 + int(ch * 0.52), text='待抽签',
                    font=rolling_font, fill=SC_MUTED,
                    state='hidden' if
                    (use_placeholder or (rolling and not use_digits)) else 'normal')
                card['blur_img'] = None
                if use_placeholder and placeholder_photo is not None:
                    card['blur_img'] = c.create_image(
                        (x1 + x2) // 2, y1 + int(ch * 0.52),
                        image=placeholder_photo)
                    self._blur_refs.append(placeholder_photo)
                card['fx'] = []
                card['fx_kind'] = None
                ccx = (x1 + x2) // 2
                ccy = y1 + int(ch * 0.52)
                if rolling and use_digits:
                    # 数字滚轮：名字文字项直接显示快速翻转的数字
                    card['fx_kind'] = 'digits'
                    c.itemconfig(card['name'], text='%02d' % (i + 1),
                                 fill='#eaf6ff')
                elif rolling and style == STYLE_RING:
                    # 光环脉冲：外环+内环，随帧从外向内收拢脉动；
                    # 用细线粗略描边，明显减少画布椭圆的锯齿感
                    card['fx_kind'] = 'ring'
                    rr = max(12, int(ch * 0.20))
                    w1 = max(1, int(rr * 0.07))
                    w2 = max(1, int(rr * 0.09))
                    card['_ring'] = (ccx, ccy, rr)
                    card['fx'] = [
                        c.create_oval(ccx - rr, ccy - rr, ccx + rr, ccy + rr,
                                      outline=DOT_T2, width=w1),
                        c.create_oval(ccx - int(rr * 0.6), ccy - int(rr * 0.6),
                                      ccx + int(rr * 0.6), ccy + int(rr * 0.6),
                                      outline=DOT_T1, width=w2),
                    ]
                elif rolling and style == STYLE_DOTS:
                    # 彗尾流光：圆点字号由卡宽约束，整排放进卡片避免溢出。
                    # 用一个实例级缓存按字号缓存“●”像素宽，二分找最大插得下的
                    # 字号，避免逐号缩小反复 tkfont.measure 卡顿。
                    if not hasattr(self, '_dot_ds_cache'):
                        self._dot_ds_cache = {}
                    avail = max(30, cw - 24)
                    factor = DOTS_N + 0.7 * (DOTS_N - 1)
                    s0 = max(12, int(ch * 0.30))
                    lo, hi = 10, s0
                    best = 12
                    while lo <= hi:
                        mid = (lo + hi) // 2
                        ds = self._dot_ds_cache.get(mid)
                        if ds is None:
                            ds = tkfont.Font(
                                master=c, font=font(mid)).measure('●')
                            self._dot_ds_cache[mid] = ds
                        if ds * factor <= avail:
                            best, lo = mid, mid + 1
                        else:
                            hi = mid - 1
                    dot_f = tkfont.Font(master=c, font=font(best))
                    ds = self._dot_ds_cache.get(best)
                    if ds is None:
                        ds = dot_f.measure('●')
                        self._dot_ds_cache[best] = ds
                    pad = int(ds * 0.7)
                    total = ds * DOTS_N + pad * (DOTS_N - 1)
                    sx = ccx - total / 2 + ds / 2
                    for k in range(DOTS_N):
                        idx = c.create_text(sx + k * (ds + pad), ccy,
                                            text='●', font=dot_f, fill=DOT_OFF)
                        card['fx'].append(idx)
                card['org'] = c.create_text(
                    (x1 + x2) // 2, y1 + int(ch * 0.78), text='',
                    font=font(max(10, min(18, int(ch * 0.11)))),
                    fill=SC_MUTED)
                self._cards.append(card)

    def _apply_locked(self, i, flash=True):
        """把第 i 张卡片切成"已定格"样式（增量 itemconfig，不重绘布局）。"""
        if i >= len(self._cards):
            return
        c = self.canvas
        card = self._cards[i]
        item = self.items[i]
        card['locked'] = True

        c.itemconfig(card['rect'], fill=SC_CARD_LOCK,
                     outline=SC_LINE_LOCK, width=2)
        c.itemconfig(card['no'], fill=ACCENT)

        name = item.get('name', '')
        avail_w = card['cw'] - 24
        if card.get('_fitted_for') != name:
            # 字号上限按能容纳约两行计算；同时下方 width=avail_w 触发自动换行，
            # 长单位名在可用宽度内折行，不再超出框线。
            cap = max(14, int(card['ch'] * 0.21))
            card['font'] = _fit_font(c, name, avail_w,
                                     min(44, cap), 12, 'bold')
            card['_fitted_for'] = name
        # 移除滚动占位的动效项（圆点/光环），再恢复名字文字项，展现最终结果
        fx = card.get('fx')
        if fx:
            for did in fx:
                try:
                    c.delete(did)
                except Exception:
                    pass
            card['fx'] = []
        blur = card.get('blur_img')
        if blur:
            try:
                c.delete(blur)
            except Exception:
                pass
            card['blur_img'] = None
        c.itemconfig(card['name'], text=name, state='normal',
                     width=avail_w, font=card['font'],
                     fill=self._main_col)

        multi = bool(self.record and (self.record.get('group_count') or 0) > 1)
        org = item.get('org') or \
            ('第 %d 组' % item['group'] if multi and item.get('group') else '')
        if org:
            org_f = _fit_font(c, org, card['cw'] - 20,
                              max(10, min(18, int(card['ch'] * 0.11))), 9)
            c.itemconfig(card['org'], text=org, font=org_f)
        else:
            c.itemconfig(card['org'], text='')

        if flash:
            c.itemconfig(card['rect'], outline=SC_FLASH, width=3)

            def _settle():
                try:
                    if card.get('locked'):
                        c.itemconfig(card['rect'], outline=SC_LINE_LOCK,
                                     width=2)
                except Exception:
                    pass
            self.win.after(240, _settle)

    # ---------- 动画 ----------
    def _random_name(self):
        pool = self.participants
        if not pool:
            return '——'
        return pool[random.randrange(len(pool))]['name']

    def _fit_cycle_text(self, name):
        """把轮播名字收进卡片可用宽度（等价旧版网页卡片的 overflow:hidden）。

        超长时逐字截断并补省略号；结果按名字缓存，65ms 换帧时零测量开销。
        """
        name = name or '——'
        fnt, max_w = self._rolling_font, self._cycle_w
        if fnt is None or max_w <= 0:
            return name
        cached = self._cycle_fit_cache.get(name)
        if cached is not None:
            return cached
        text = name
        try:
            if fnt.measure(text) > max_w:
                ell = '…'
                ell_w = fnt.measure(ell)
                keep = len(text)
                while keep > 0 and \
                        fnt.measure(text[:keep]) + ell_w > max_w:
                    keep -= 1
                text = (text[:keep] + ell) if keep > 0 else ell
        except Exception:
            text = name
        self._cycle_fit_cache[name] = text
        return text

    def start_draw(self):
        if self.mode != 'idle':
            return
        record = self.admin.perform_draw_record()
        if record is None:
            self.toast('至少需要 2 名参赛者才能抽签')
            return
        # 按后台设置的速度倍率缩放各阶段时长（倍率越大越快）
        sp = max(0.3, float(
            self.admin.state.get('settings', {}).get('speed') or 1.0))
        self._roll_ms = max(300, int(ROLL_MS / sp))
        self._cycle_ms = max(18, int(CYCLE_MS / sp))
        self._auto_total_ms = max(800, int(AUTO_TOTAL_MS / sp))
        self.record = record
        self.items = record.get('items') or []
        self.reveal = 0
        self.auto = True
        self._rolling_names = ['待抽签'] * len(self.items)
        self.mode = 'rolling'
        self._render()
        self._start_cycling()
        if self.music_on:
            path = self.admin.get_music_path()
            if path:
                self._mci.play(path)
        self._step_job = self.win.after(self._roll_ms, self._jump_to_locking)

    def _start_cycling(self):
        self._stop_cycling()

        def cycle():
            if self._closed or self.mode != 'rolling':
                return
            try:
                if not self.win.winfo_exists():
                    return
            except Exception:
                return
            # 只 itemconfig 未定格卡片的姓名，不动布局
            if len(self._cards) < len(self.items):
                self._render()
                return
            c = self.canvas
            t = self._roll_tick
            for i in range(len(self.items)):
                if i >= self.reveal:
                    card = self._cards[i]
                    kind = card.get('fx_kind')
                    if kind == 'dots':      # 彗尾流光
                        fx = card['fx']
                        if fx:
                            head = (t + i) % DOTS_N
                            for k, did in enumerate(fx):
                                d = (head - k) % DOTS_N
                                c.itemconfig(
                                    did, fill=DOT_COLORS[d] if d < len(DOT_COLORS)
                                    else DOT_OFF)
                    elif kind == 'ring':    # 光环脉冲：双环周期收拢
                        fx = card['fx']
                        if len(fx) >= 2:
                            cx, cy, rr = card['_ring']
                            ph = (t + i) % 9
                            sc = 1.0 - ph / 9.0 * 0.6   # 0.4 → 1.0
                            r2 = rr * sc
                            r1 = rr * 0.6 * (2.0 - sc)
                            c.coords(fx[1], cx - r2, cy - r2, cx + r2, cy + r2)
                            c.coords(fx[0], cx - r1, cy - r1, cx + r1, cy + r1)
                    elif kind == 'digits':  # 数字滚轮
                        n = len(self.items)
                        c.itemconfig(card['name'],
                                     text='%02d' % ((t * 7 + i * 3) % max(n, 1) + 1))
            self._roll_tick += 1
            self._cycle_job = self.win.after(self._cycle_ms, cycle)

        self._cycle_job = self.win.after(self._cycle_ms, cycle)

    def _stop_cycling(self):
        if self._cycle_job:
            try:
                self.win.after_cancel(self._cycle_job)
            except Exception:
                pass
            self._cycle_job = None

    def _jump_to_locking(self):
        if self.mode != 'rolling':
            return
        # 不停止循环动画：进入逐张揭示后，未揭晓卡片继续滚动
        # 直到各自定格，避免"结果未出动画先停"
        self.mode = 'locking'
        self._refresh_chrome()
        self._start_auto_reveal()

    def _start_auto_reveal(self):
        total = len(self.items)
        per = max(80, min(1200, self._auto_total_ms // max(total, 1)))

        def step():
            if self._closed or self.mode != 'locking':
                return
            if self.reveal >= total:
                self._finish()
                return
            self._reveal_one()
            if self.auto and self.reveal < total:
                self._step_job = self.win.after(per, step)
            else:
                self._finish()

        step()

    def _reveal_one(self):
        if self.reveal >= len(self.items):
            return
        i = self.reveal
        self.reveal += 1
        if i < len(self._cards) and i < len(self.items):
            self._apply_locked(i)
        else:
            self._render()
        if self.sound_on:
            blip()

    def _finish(self):
        self._stop_cycling()
        self.mode = 'done'
        self._refresh_chrome()
        if self.sound_on:
            fanfare()
        self._burst()

    def advance(self):
        if self.mode == 'idle':
            self.start_draw()
        elif self.mode == 'rolling':
            self._jump_to_locking()
        elif self.mode == 'locking':
            self.auto = False
            if self._step_job:
                try:
                    self.win.after_cancel(self._step_job)
                except Exception:
                    pass
                self._step_job = None
            if self.reveal < len(self.items):
                self._reveal_one()
            else:
                self._finish()
        elif self.mode == 'done':
            # 抽签全程只允许一次：大屏不能自行重抽，必须回管理后台重置
            self.toast('抽签已完成 · 如需重抽请回管理后台点「重置结果」')

    def on_admin_reset(self):
        """管理后台清除结果后调用：大屏同步回到待抽签状态（不在此回调后台）。"""
        self._stop_cycling()
        if self._step_job:
            try:
                self.win.after_cancel(self._step_job)
            except Exception:
                pass
            self._step_job = None
        self._clear_confetti()
        self._mci.stop()
        self.mode = 'idle'
        self.items = []
        self.record = None
        self.reveal = 0
        self._render()

    # ---------- 庆祝特效（双侧礼花炮 + 中心金色光环） ----------
    FX_COLORS = ('#f6c453', '#ffd54f', '#ffe9a8', '#ffffff',
                 '#5ad1e6', '#8fd4ff', '#ff9ecb')
    FX_GOLD = ('#f6c453', '#ffd54f', '#ffe9a8', '#ffffff')
    FX_RING_COLORS = ('#ffd54f', '#ffe9a8', '#f6c453')
    FX_FRAME_MS = 30

    def _schedule_fx(self, delay, fn):
        """延迟任务统一登记，特效清除/窗口关闭时可整体取消。"""
        job = []

        def wrapper():
            try:
                self._fx_wave_jobs.remove(job[0])
            except (ValueError, AttributeError):
                pass
            fn()

        job.append(self.win.after(delay, wrapper))
        self._fx_wave_jobs.append(job[0])

    def _burst(self):
        """全部揭晓时的庆祝：左右下角礼花炮分 4 波 + 中心金色三环脉冲。"""
        try:
            w = self.canvas.winfo_width()
            h = self.canvas.winfo_height()
        except Exception:
            return
        if w < 100 or h < 100:
            return
        self._clear_confetti()
        # 首发较密形成高潮，后续三波补射，持续约 1 秒
        for delay, count in ((0, 15), (300, 11), (620, 11), (940, 9)):
            self._schedule_fx(delay, lambda n=count: self._fire_cannons(n))
        for i in range(3):
            self._schedule_fx(i * 140, lambda k=i: self._spawn_ring(k))
        if self._confetti_job is None:
            self._confetti_job = self.win.after(self.FX_FRAME_MS,
                                                self._fx_tick)

    def _fire_cannons(self, count):
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        for _ in range(count):
            self._confetti.append(self._make_particle(w, h, 1))
            self._confetti.append(self._make_particle(w, h, -1))

    def _make_particle(self, w, h, side):
        """side=+1 左下角发射（向右上），side=-1 右下角发射（向左上）。"""
        kind = random.choices(('stream', 'petal', 'dot'),
                              weights=(5, 3, 2))[0]
        gold = kind != 'dot' and random.random() < 0.55
        sf = min(1.5, max(0.9, h / 1080.0))  # 大屏粒子尺寸等比放大
        p = {
            'kind': kind,
            'x': w * (0.015 if side > 0 else 0.985),
            'y': h * 0.99,
            'vx': side * (7 + random.random() * 6)
                  + (random.random() - 0.5) * 4,
            'vy': -(17 + random.random() * 11),  # 高仰角：抛物线越过卡片区
            'angle': random.random() * math.tau,
            'vr': (random.random() - 0.5) * 0.42,
            'phase': random.random() * math.tau,
            'sway_spd': 0.12 + random.random() * 0.14,
            'sway': (0.5 + random.random() * 1.1) * sf,
            'color': random.choice(self.FX_GOLD if gold else self.FX_COLORS),
            'item': None}
        if kind == 'stream':  # 长条彩带
            hw_ = (3 + random.random() * 2) * sf
            hh = (6.5 + random.random() * 3.5) * sf
            p['pts'] = [(-hw_, -hh), (hw_, -hh), (hw_, hh), (-hw_, hh)]
        elif kind == 'petal':  # 花瓣（叶形菱形）
            r = (4.5 + random.random() * 3.5) * sf
            p['pts'] = [(0, -r), (r * 0.62, 0), (0, r), (-r * 0.62, 0)]
        else:                 # 圆点
            p['r'] = (2.8 + random.random() * 2.4) * sf
        return p

    def _rotated_points(self, p):
        ca, sa = math.cos(p['angle']), math.sin(p['angle'])
        out = []
        for px, py in p['pts']:
            out.append(p['x'] + px * ca - py * sa)
            out.append(p['y'] + px * sa + py * ca)
        return out

    def _create_particle_item(self, canvas, p):
        if p['kind'] == 'dot':
            r = p['r']
            p['item'] = canvas.create_oval(
                p['x'] - r, p['y'] - r, p['x'] + r, p['y'] + r,
                fill=p['color'], width=0)
        else:
            p['item'] = canvas.create_polygon(
                *self._rotated_points(p), fill=p['color'], width=0)

    def _spawn_ring(self, idx):
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        self._rings.append({
            'cx': w // 2, 'cy': int(h * 0.44),
            'r0': max(28, int(h * 0.05)),
            'r1': int(min(w, h) * 0.40),
            'dur': 30, 't': 0,
            'w0': 5 if idx == 0 else 4,
            'c0': self.FX_RING_COLORS[idx % len(self.FX_RING_COLORS)],
            'item': None})

    def _fx_tick(self):
        """统一帧循环：粒子物理（重力/阻力/飘摆/旋转）+ 光环扩散。"""
        self._confetti_job = None
        if self._closed:
            return
        try:
            if not self.win.winfo_exists():
                return
        except Exception:
            return
        c = self.canvas
        w, h = c.winfo_width(), c.winfo_height()

        keep = []
        for p in self._confetti:
            p['vy'] = min(p['vy'] + 0.30, 15)
            p['vx'] *= 0.992
            p['phase'] += p['sway_spd']
            p['x'] += p['vx'] + math.sin(p['phase']) * p['sway']
            p['y'] += p['vy']
            p['angle'] += p['vr']
            # 飞出屏幕即移除（不做整屏瞬隐，杜绝"半空突然消失"）
            if (p['y'] > h + 40 or p['x'] < -60 or p['x'] > w + 60
                    or p['y'] < -h):
                if p['item'] is not None:
                    try:
                        c.delete(p['item'])
                    except Exception:
                        pass
                continue
            try:
                if p['item'] is None:
                    self._create_particle_item(c, p)
                elif p['kind'] == 'dot':
                    r = p['r']
                    c.coords(p['item'], p['x'] - r, p['y'] - r,
                             p['x'] + r, p['y'] + r)
                else:
                    c.coords(p['item'], *self._rotated_points(p))
            except Exception:
                pass
            keep.append(p)
        self._confetti = keep

        rkeep = []
        for ring in self._rings:
            ring['t'] += 1
            t = ring['t'] / ring['dur']
            if t >= 1:
                if ring['item'] is not None:
                    try:
                        c.delete(ring['item'])
                    except Exception:
                        pass
                continue
            e = 1 - (1 - t) ** 2  # ease-out
            r = ring['r0'] + (ring['r1'] - ring['r0']) * e
            col = ui_theme.mix(ring['c0'], SC_BG, t)  # 描边色渐隐入背景
            wid = max(1, ring['w0'] * (1 - t))
            cx, cy = ring['cx'], ring['cy']
            try:
                if ring['item'] is None:
                    ring['item'] = c.create_oval(
                        cx - r, cy - r, cx + r, cy + r,
                        outline=col, width=wid)
                else:
                    c.coords(ring['item'], cx - r, cy - r, cx + r, cy + r)
                    c.itemconfig(ring['item'], outline=col, width=wid)
            except Exception:
                pass
            rkeep.append(ring)
        self._rings = rkeep

        if self._confetti or self._rings or self._fx_wave_jobs:
            self._confetti_job = self.win.after(self.FX_FRAME_MS,
                                                self._fx_tick)

    def _redraw_fx(self, canvas, w, h):
        """整屏重绘（delete all）后按当前状态重建特效画布项。"""
        for p in self._confetti:
            p['item'] = None
            try:
                self._create_particle_item(canvas, p)
            except Exception:
                pass
        for ring in self._rings:
            ring['item'] = None  # 下一帧按当前进度重建
        if (self._confetti or self._rings) and self._confetti_job is None:
            self._confetti_job = self.win.after(self.FX_FRAME_MS,
                                                self._fx_tick)

    def _clear_confetti(self):
        for p in self._confetti:
            if p.get('item') is not None:
                try:
                    self.canvas.delete(p['item'])
                except Exception:
                    pass
        for ring in self._rings:
            if ring.get('item') is not None:
                try:
                    self.canvas.delete(ring['item'])
                except Exception:
                    pass
        self._confetti = []
        self._rings = []
        for job in list(self._fx_wave_jobs):
            try:
                self.win.after_cancel(job)
            except Exception:
                pass
        self._fx_wave_jobs = []
        if self._confetti_job:
            try:
                self.win.after_cancel(self._confetti_job)
            except Exception:
                pass
            self._confetti_job = None

    # ---------- 轻提示 ----------
    def toast(self, message):
        """显示轻提示（2.2s 后自动消失）。message 存到 self._toast_msg，
        以便 _render 整屏重绘（delete all）后自动恢复，不被抹掉。"""
        self._toast_msg = message
        self._draw_toast(message)
        if self._toast_job:
            try:
                self.win.after_cancel(self._toast_job)
            except Exception:
                pass
        self._toast_job = self.win.after(2200, lambda: self._hide_toast())

    def _draw_toast(self, message):
        """只负责在画布上绘制 toast 底托 + 文字（不管理计时）。"""
        try:
            self.canvas.delete('toast')
            w = self.canvas.winfo_width()
            h = self.canvas.winfo_height()
            # 标题 + 副标题 + 元信息约占屏高 13%，卡片从 17% 起，
            # toast 放在两者之间的空隙，避免压标题也不挡卡片
            y = max(110, int(h * 0.145))
            f = tkfont.Font(master=self.canvas, font=font(15, 'bold'))
            tw = max(120, f.measure(message))
            pad_x, pad_y = 22, 10
            # 半透明深色底托，保证在任何背景上都清晰可读
            try:
                from PIL import Image, ImageTk
                from PIL import ImageDraw
                bw, bh = tw + pad_x * 2, f.metrics('linespace') + pad_y * 2
                img = Image.new('RGBA', (bw, bh), (0, 0, 0, 0))
                d = ImageDraw.Draw(img)
                d.rounded_rectangle((0, 0, bw - 1, bh - 1), radius=12,
                                    fill=(8, 15, 35, 210))
                self._toast_img = ImageTk.PhotoImage(img)
                self.canvas.create_image(w // 2, y, image=self._toast_img,
                                         tags='toast')
            except Exception:
                pass
            self.canvas.create_text(w // 2, y, text=message, tags='toast',
                                    font=f, fill=SC_TEXT)
        except Exception:
            pass

    def _hide_toast(self):
        self._toast_msg = None
        self._toast_job = None
        try:
            self.canvas.delete('toast')
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 管理窗口
# ---------------------------------------------------------------------------
class LotteryAdmin(object):
    """管理窗口：设置 / 名单 / 抽签 / 结果 / 历史 / 校验 / 导出 / 音乐。"""

    def __init__(self, win, on_close=None):
        self.win = win
        self.on_close = on_close      # 会务助手注入：统一走 _close_lottery_win
        self.state = load_state()
        self.screen = None
        self._save_job = None
        self._viewing = None
        self._preview = MciPlayer('lottery_prev')

        ui_theme.style_window(win)
        win.title('现场抽签排序')
        win.configure(bg=PALETTE['bg'])

        band = GradientBand(win, title='现场抽签排序',
                            subtitle='可复现 · 可审计的出场顺序抽签',
                            color_from=ACCENT_DARK, color_to=ACCENT, height=54)
        band.pack(fill='x')

        body = tk.Frame(win, bg=PALETTE['bg'])
        body.pack(fill='both', expand=True, padx=14, pady=12)

        # ---------- 左列 ----------
        left = tk.Frame(body, bg=PALETTE['bg'])
        left.pack(side='left', fill='y', padx=(0, 12))
        self._build_settings(left)
        self._build_participants(left)
        self._build_music(left)

        # ---------- 右列 ----------
        right = tk.Frame(body, bg=PALETTE['bg'])
        right.pack(side='left', fill='both', expand=True)
        self._build_actions(right)
        self._build_result(right)
        self._build_records(right)

        win.protocol('WM_DELETE_WINDOW', self.close_requested)
        self.refresh_all()

    # ---------- 设置 ----------
    def _build_settings(self, parent):
        box = ttk.Labelframe(parent, text='会议设置', padding=10)
        box.pack(fill='x')

        row = tk.Frame(box, bg=PALETTE['bg'])
        row.pack(fill='x', pady=2)
        tk.Label(row, text='标题', bg=PALETTE['bg'], fg=PALETTE['muted'],
                 font=font(10), width=6, anchor='w').pack(side='left')
        self.var_title = tk.StringVar(value=self.state['settings'].get('title', ''))
        ttk.Entry(row, textvariable=self.var_title).pack(
            side='left', fill='x', expand=True, padx=(0, 4))

        row = tk.Frame(box, bg=PALETTE['bg'])
        row.pack(fill='x', pady=2)
        tk.Label(row, text='副标题', bg=PALETTE['bg'], fg=PALETTE['muted'],
                 font=font(10), width=6, anchor='w').pack(side='left')
        self.var_subtitle = tk.StringVar(
            value=self.state['settings'].get('subtitle', ''))
        ttk.Entry(row, textvariable=self.var_subtitle).pack(
            side='left', fill='x', expand=True, padx=(0, 4))

        row = tk.Frame(box, bg=PALETTE['bg'])
        row.pack(fill='x', pady=(2, 0))
        tk.Label(row, text='分组数', bg=PALETTE['bg'], fg=PALETTE['muted'],
                 font=font(10), width=6, anchor='w').pack(side='left')
        self.var_groups = tk.IntVar(
            value=int(self.state['settings'].get('group_count') or 0))
        ttk.Spinbox(row, from_=0, to=20, textvariable=self.var_groups,
                    width=6).pack(side='left', padx=(0, 8))
        tk.Label(row, text='抽签速度', bg=PALETTE['bg'], fg=PALETTE['muted'],
                 font=font(10), width=8, anchor='w').pack(side='left', padx=(8, 0))
        cur_speed = float(self.state['settings'].get('speed') or 1.0)
        # 反查当前倍率对应的档位名，找不到则取最接近的
        names = list(SPEED_PRESETS.keys())
        best = min(names, key=lambda n: abs(SPEED_PRESETS[n] - cur_speed))
        self.var_speed = tk.StringVar(value=best)
        ttk.Combobox(row, textvariable=self.var_speed, values=names,
                     state='readonly', width=6).pack(side='left', padx=(0, 8))
        tk.Label(row, text='0 = 不分组', bg=PALETTE['bg'],
                 fg=PALETTE['faint'], font=font(9)).pack(side='left', padx=8)

        row = tk.Frame(box, bg=PALETTE['bg'])
        row.pack(fill='x', pady=(6, 0))
        tk.Label(row, text='抽签样式', bg=PALETTE['bg'], fg=PALETTE['muted'],
                 font=font(10), width=6, anchor='w').pack(side='left')
        cur_style = self.state['settings'].get('screen_style') or STYLE_DEF_DEFAULT
        self.var_style = tk.StringVar(
            value=cur_style if cur_style in STYLE_OPTIONS else STYLE_DEF_DEFAULT)
        ttk.Combobox(row, textvariable=self.var_style, values=STYLE_OPTIONS,
                     state='readonly', width=8).pack(side='left', padx=(0, 8))
        FlatButton(row, text='选择背景图', kind='plain', command=self.choose_bg).pack(
            side='left', padx=(0, 6))
        FlatButton(row, text='清除背景', kind='danger',
                   command=self.clear_bg).pack(side='left')
        self.lbl_bg = tk.Label(row, text='', bg=PALETTE['bg'],
                              fg=PALETTE['primary'], font=font(9))
        self.lbl_bg.pack(side='left', padx=(8, 0))
        self._refresh_bg_label()

        # 字体 + 文字颜色
        row = tk.Frame(box, bg=PALETTE['bg'])
        row.pack(fill='x', pady=(6, 0))
        tk.Label(row, text='大屏字体', bg=PALETTE['bg'], fg=PALETTE['muted'],
                 font=font(10), width=6, anchor='w').pack(side='left')
        self.var_font = tk.StringVar(
            value=self.state['settings'].get('font_family') or '')
        fams = ui_theme_families()
        ttk.Combobox(row, textvariable=self.var_font, values=fams,
                     width=13).pack(side='left', padx=(0, 10))
        tk.Label(row, text='文字颜色', bg=PALETTE['bg'], fg=PALETTE['muted'],
                 font=font(10)).pack(side='left', padx=(0, 4))
        cur_col = self.state['settings'].get('text_color') or '#ffffff'
        self.var_color = tk.StringVar(
            value=cur_col if _is_hex(cur_col) else '#ffffff')
        self.swatch = tk.Label(row, text='    ', bg=self.var_color.get(),
                              relief='solid', bd=1)
        self.swatch.pack(side='left', padx=(0, 4))
        FlatButton(row, text='选色', kind='plain', command=self.pick_text_color).pack(
            side='left')

        # Logo 上传
        row = tk.Frame(box, bg=PALETTE['bg'])
        row.pack(fill='x', pady=(6, 0))
        tk.Label(row, text='Logo', bg=PALETTE['bg'], fg=PALETTE['muted'],
                 font=font(10), width=6, anchor='w').pack(side='left')
        FlatButton(row, text='上传Logo', kind='plain', command=self.choose_logo).pack(
            side='left', padx=(0, 6))
        FlatButton(row, text='清除Logo', kind='danger',
                   command=self.clear_logo).pack(side='left')
        self.lbl_logo = tk.Label(row, text='', bg=PALETTE['bg'],
                                 fg=PALETTE['primary'], font=font(9))
        self.lbl_logo.pack(side='left', padx=(8, 0))
        self._refresh_logo_label()

        # 保存设置：放在所有设置项底部、居右醒目，避免上面改完还要回顶保存
        row = tk.Frame(box, bg=PALETTE['bg'])
        row.pack(fill='x', pady=(8, 0))
        tk.Frame(row, bg=PALETTE['bg']).pack(fill='x', expand=True, side='left')
        self.save_btn = FlatButton(row, text='保存设置', kind='primary',
                                   command=self.save_settings)
        self.save_btn.pack(side='right')

    def save_settings(self):
        try:
            groups = max(0, int(self.var_groups.get()))
        except (tk.TclError, ValueError):
            groups = 0
        self.state['settings']['title'] = self.var_title.get().strip()
        self.state['settings']['subtitle'] = self.var_subtitle.get().strip()
        self.state['settings']['group_count'] = groups
        self.state['settings']['speed'] = SPEED_PRESETS.get(
            self.var_speed.get(), 1.0)
        self.state['settings']['screen_style'] = \
            self.var_style.get() if self.var_style.get() in STYLE_OPTIONS \
            else STYLE_DEF_DEFAULT
        self.state['settings']['font_family'] = \
            (self.var_font.get() or '').strip()
        self.state['settings']['text_color'] = \
            self.var_color.get() if _is_hex(self.var_color.get()) else '#ffffff'
        save_state(self.state)
        self._status('设置已保存')
        self._refresh_screen()

    def pick_text_color(self):
        from tkinter import colorchooser
        cur = self.var_color.get()
        try:
            rgb, hx = colorchooser.askcolor(
                color=cur, title='选择大屏文字颜色',
                parent=self.win)
        except Exception:
            return
        if hx and _is_hex(hx):
            self.var_color.set(hx)
            try:
                self.swatch.config(bg=hx)
            except Exception:
                pass
            self.state['settings']['text_color'] = hx
            save_state(self.state)
            self._refresh_screen()

    def choose_logo(self):
        self._keep_front()
        p = filedialog.askopenfilename(
            title='选择 Logo 图片', parent=self.win,
            filetypes=[('图片', '*.png *.jpg *.jpeg *.bmp'),
                       ('所有文件', '*.*')])
        if not p:
            return
        try:
            from PIL import Image
            Image.open(p).verify()
        except Exception:
            messagebox.showwarning('Logo', '无法读取该图片，请重新选择。')
            return
        self.state['settings']['logo_path'] = p
        save_state(self.state)
        self._refresh_logo_label('Logo 已上传')
        self._refresh_screen()

    def clear_logo(self):
        if self.state['settings'].get('logo_path'):
            self.state['settings']['logo_path'] = ''
            save_state(self.state)
            self._refresh_logo_label('Logo 已清除')
            self._refresh_screen()

    def _refresh_logo_label(self, hint=None):
        path = self.state['settings'].get('logo_path') or ''
        text = hint if hint is not None else \
            (('已上传: ' + Path(path).name) if path else '未上传')
        try:
            self.lbl_logo.config(text=text)
        except Exception:
            pass

    def choose_bg(self):
        self._keep_front()
        p = filedialog.askopenfilename(
            title='选择大屏背景图', parent=self.win,
            filetypes=[('图片', '*.png *.jpg *.jpeg *.bmp *.gif'),
                       ('所有文件', '*.*')])
        if not p:
            return
        try:
            from PIL import Image
            Image.open(p).verify()
        except Exception:
            messagebox.showwarning('背景图', '无法读取该图片，请重新选择。')
            return
        self.state['settings']['screen_bg'] = p
        save_state(self.state)
        self._refresh_bg_label('背景图已设置')
        self._refresh_screen()

    def clear_bg(self):
        if self.state['settings'].get('screen_bg'):
            self.state['settings']['screen_bg'] = ''
            save_state(self.state)
            self._refresh_bg_label('背景已清除')
            self._refresh_screen()

    def _refresh_bg_label(self, hint=None):
        bg = self.state['settings'].get('screen_bg') or ''
        text = hint if hint is not None else \
            (('已设置: ' + Path(bg).name) if bg else '未设置')
        try:
            self.lbl_bg.config(text=text)
        except Exception:
            pass

    def _refresh_screen(self):
        scr = getattr(self, 'screen', None)
        if scr and scr.is_open():
            try:
                scr._render()
            except Exception:
                pass

    def _keep_front(self):
        # 交互期间把后台窗口临时置顶，压住其它置顶小窗（如首页弹窗），
        # 一段时间后自动解除，避免一直霸占顶层
        try:
            w = self.win.winfo_toplevel()
            w.attributes('-topmost', True)
            w.lift()
            w.focus_force()
            self.win.after(2500, lambda: _safe_no_topmost(w))
        except Exception:
            pass

    # ---------- 名单 ----------
    def _build_participants(self, parent):
        box = ttk.Labelframe(parent, text='参赛名单', padding=8)
        box.pack(fill='both', expand=True, pady=(10, 0))

        bar = tk.Frame(box, bg=PALETTE['bg'])
        bar.pack(fill='x', pady=(0, 6))
        FlatButton(bar, text='导入名单', command=self.import_file).pack(
            side='left')
        FlatButton(bar, text='粘贴导入', command=self.paste_import).pack(
            side='left', padx=6)
        FlatButton(bar, text='添加', command=self.add_row).pack(side='left')
        FlatButton(bar, text='清空', kind='danger',
                   command=self.clear_list).pack(side='left', padx=6)
        self.lbl_count = tk.Label(bar, text='', bg=PALETTE['bg'],
                                  fg=PALETTE['muted'], font=font(9))
        self.lbl_count.pack(side='right')

        # 表头
        head = tk.Frame(box, bg=PALETTE['bg'])
        head.pack(fill='x')
        tk.Label(head, text='序号', bg=PALETTE['bg'], fg=PALETTE['muted'],
                 font=font(9), width=4, anchor='center').pack(
            side='left', padx=(2, 0))
        tk.Label(head, text='姓名', bg=PALETTE['bg'], fg=PALETTE['muted'],
                 font=font(9), width=10, anchor='w').pack(side='left', padx=4)
        tk.Label(head, text='单位 / 部门', bg=PALETTE['bg'],
                 fg=PALETTE['muted'], font=font(9), width=12,
                 anchor='w').pack(side='left', padx=4)
        tk.Label(head, text='备注', bg=PALETTE['bg'], fg=PALETTE['muted'],
                 font=font(9), width=10, anchor='w').pack(side='left', padx=4)
        tk.Label(head, text='', bg=PALETTE['bg'], width=3).pack(side='left')

        # 可滚动行区
        wrap = tk.Frame(box, bg=PALETTE['surface'])
        wrap.pack(fill='both', expand=True)
        self.rows_canvas = tk.Canvas(wrap, bg=PALETTE['surface'],
                                     highlightthickness=0, bd=0)
        scroll = ttk.Scrollbar(wrap, orient='vertical',
                               command=self.rows_canvas.yview)
        self.rows_canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')
        self.rows_canvas.pack(side='left', fill='both', expand=True)
        self.rows_frame = tk.Frame(self.rows_canvas, bg=PALETTE['surface'])
        self.rows_win = self.rows_canvas.create_window(
            (0, 0), window=self.rows_frame, anchor='nw')
        self.rows_frame.bind(
            '<Configure>',
            lambda e: self.rows_canvas.configure(
                scrollregion=self.rows_canvas.bbox('all')))
        self.rows_canvas.bind(
            '<Configure>',
            lambda e: self.rows_canvas.itemconfigure(self.rows_win,
                                                      width=e.width))
        # 鼠标滚轮：只在名单区上方生效（bind_all 会劫持整个程序的滚轮）
        def _wheel(e):
            try:
                if self.rows_canvas.winfo_exists():
                    self.rows_canvas.yview_scroll(
                        int(-1 * (e.delta / 120)), 'units')
            except Exception:
                pass

        def _bind_wheel(_e=None):
            self.rows_canvas.bind_all('<MouseWheel>', _wheel)

        def _unbind_wheel(_e=None):
            self.rows_canvas.unbind_all('<MouseWheel>')

        wrap.bind('<Enter>', _bind_wheel)
        wrap.bind('<Leave>', _unbind_wheel)

    def rebuild_rows(self):
        try:
            for child in self.rows_frame.winfo_children():
                child.destroy()
        except Exception:
            pass
        parts = self.state['participants']
        for i, p in enumerate(parts):
            row = tk.Frame(self.rows_frame, bg=PALETTE['surface'])
            row.pack(fill='x', pady=1)
            tk.Label(row, text=str(i + 1), bg=PALETTE['surface'],
                     fg=PALETTE['faint'], font=font(9), width=4,
                     anchor='center').pack(side='left', padx=(2, 0))

            for field, width in (('name', 12), ('org', 16), ('note', 14)):
                var = tk.StringVar(value=p.get(field) or '')

                def _make_trace(idx=i, fld=field, v=var):
                    def _trace(*_a):
                        try:
                            self.state['participants'][idx][fld] = v.get().strip()
                            self.schedule_save()
                        except (IndexError, tk.TclError):
                            pass
                    return _trace

                var.trace('w', _make_trace())
                ent = ttk.Entry(row, textvariable=var, width=width)
                ent.pack(side='left', padx=4, ipady=1)

            def _make_del(idx=i):
                return lambda: self.delete_row(idx)

            btn = tk.Label(row, text='×', bg=PALETTE['surface'],
                           fg=PALETTE['faint'], font=font(11, 'bold'),
                           width=3, cursor='hand2')
            btn.bind('<Button-1>', lambda e, fn=_make_del(): fn())
            btn.pack(side='left')

        self.lbl_count.configure(text='共 %d 人' % len(
            _clean_participants(parts)))
        self.rows_canvas.update_idletasks()
        self.rows_canvas.configure(scrollregion=self.rows_canvas.bbox('all'))

    def schedule_save(self):
        if self._save_job:
            try:
                self.win.after_cancel(self._save_job)
            except Exception:
                pass
        self._save_job = self.win.after(700, self._flush_pending)

    def _flush_pending(self):
        self._save_job = None
        if self.state['participants'] != self._last_saved_parts:
            self.state['participants'] = _clean_participants(
                self.state['participants'])
            save_state(self.state)
            self._last_saved_parts = json.loads(json.dumps(
                self.state['participants']))
            self.lbl_count.configure(
                text='共 %d 人' % len(self.state['participants']))

    def import_file(self):
        path = filedialog.askopenfilename(
            parent=self.win, title='选择名单文件',
            filetypes=[('名单文件', '*.xlsx *.xlsm *.csv *.txt'),
                       ('Excel', '*.xlsx *.xlsm'), ('CSV/文本', '*.csv *.txt'),
                       ('所有文件', '*.*')])
        if not path:
            return
        try:
            rows = parse_upload_file(path)
        except ValueError as e:
            messagebox.showerror('导入失败', str(e), parent=self.win)
            return
        except Exception as e:
            messagebox.showerror('导入失败', '解析失败：%s' % e, parent=self.win)
            return
        if not rows:
            messagebox.showwarning('导入失败', '没有解析到有效名单，请检查文件内容',
                                   parent=self.win)
            return
        self._flush_pending()
        self.state['participants'] = rows
        self.state['result'] = None
        save_state(self.state)
        self._last_saved_parts = json.loads(json.dumps(rows))
        self.rebuild_rows()
        self.refresh_result()
        self._status('已导入 %d 人' % len(rows))

    def paste_import(self):
        dlg = tk.Toplevel(self.win)
        dlg.title('粘贴导入名单')
        dlg.configure(bg=PALETTE['bg'])
        dlg.transient(self.win)
        ui_theme.style_window(dlg)

        # 按钮条先 pack 到最底部并保留位置，避免 Text 把空间吃光后
        # 按钮被挤出可视区域（之前的“粘贴后看不到确定按钮”）
        bar = tk.Frame(dlg, bg=PALETTE['bg'])
        bar.pack(side='bottom', fill='x', padx=12, pady=(8, 12))

        tk.Label(dlg, bg=PALETTE['bg'], fg=PALETTE['muted'], font=font(10),
                 text='每行一人：姓名[,单位[,备注]]，支持逗号 / 制表符分隔\n'
                      '也可直接粘贴 Excel 区域').pack(anchor='w', padx=12,
                                                      pady=(10, 4))
        txt = tk.Text(dlg, bg=PALETTE['surface'], fg=PALETTE['text'],
                      font=font(10), height=14, undo=True, relief='flat',
                      highlightthickness=1,
                      highlightbackground=PALETTE['border'])
        txt.pack(fill='both', expand=True, padx=12)

        def _apply(mode):
            rows = parse_pasted_text(txt.get('1.0', 'end'))
            if not rows:
                messagebox.showwarning('粘贴导入', '没有解析到有效名单',
                                       parent=dlg)
                return
            self._flush_pending()
            if mode == 'append':
                merged = self.state['participants'] + rows
            else:
                merged = rows
            self.state['participants'] = merged
            self.state['result'] = None
            save_state(self.state)
            self._last_saved_parts = json.loads(json.dumps(merged))
            self.rebuild_rows()
            self.refresh_result()
            self._status('已%s %d 人，现共 %d 人'
                         % ('追加' if mode == 'append' else '导入',
                            len(rows), len(merged)))
            dlg.destroy()

        FlatButton(bar, text='替换导入', kind='primary',
                   command=lambda: _apply('replace')).pack(side='left')
        FlatButton(bar, text='追加导入',
                   command=lambda: _apply('append')).pack(side='left', padx=8)
        FlatButton(bar, text='取消', command=dlg.destroy).pack(side='right')

        # 控件全部创建后再按实际尺寸相对管理窗口居中
        dlg.update_idletasks()
        w, h = 500, 420
        px = self.win.winfo_rootx() + self.win.winfo_width() // 2
        py = self.win.winfo_rooty() + self.win.winfo_height() // 2
        x = px - w // 2
        y = py - h // 2
        dlg.geometry('%dx%d+%d+%d' % (w, h, max(0, x), max(0, y)))
        # Windows 下 geometry 坐标含标题栏外框、winfo_root* 是客户区
        # 原点，先落一次再按实测中心校正，保证视觉上真正居中
        dlg.update_idletasks()
        x += px - (dlg.winfo_rootx() + dlg.winfo_width() // 2)
        y += py - (dlg.winfo_rooty() + dlg.winfo_height() // 2)
        dlg.geometry('+%d+%d' % (max(0, x), max(0, y)))
        dlg.minsize(460, 360)
        dlg.grab_set()
        txt.focus_set()

    def add_row(self):
        self._flush_pending()
        self.state['participants'].append({'name': '', 'org': '', 'note': ''})
        self.rebuild_rows()
        self.schedule_save()

    def delete_row(self, idx):
        try:
            del self.state['participants'][idx]
        except (IndexError, tk.TclError):
            return
        self.rebuild_rows()
        self.schedule_save()

    def clear_list(self):
        if not self.state['participants']:
            return
        if not messagebox.askyesno(
                '清空名单',
                '确定清空全部 %d 名参赛者？' % len(self.state['participants']),
                parent=self.win):
            return
        self.state['participants'] = []
        save_state(self.state)
        self._last_saved_parts = []
        self.rebuild_rows()
        self._status('名单已清空')

    # ---------- 背景音乐 ----------
    def _build_music(self, parent):
        box = ttk.Labelframe(parent, text='背景音乐', padding=8)
        box.pack(fill='x', pady=(10, 0))

        bar = tk.Frame(box, bg=PALETTE['bg'])
        bar.pack(fill='x')
        FlatButton(bar, text='选择音乐', command=self.choose_music).pack(
            side='left')
        FlatButton(bar, text='试听', command=self.toggle_preview).pack(
            side='left', padx=6)
        FlatButton(bar, text='删除', kind='danger',
                   command=self.delete_music).pack(side='left')
        self.lbl_music = tk.Label(bar, text='未设置', bg=PALETTE['bg'],
                                  fg=PALETTE['muted'], font=font(9))
        self.lbl_music.pack(side='left', padx=10)

    def get_music_path(self):
        music = self.state.get('music') or {}
        name = music.get('filename')
        if not name:
            return None
        path = MUSIC_DIR / name
        return str(path) if path.exists() else None

    def choose_music(self):
        path = filedialog.askopenfilename(
            parent=self.win, title='选择背景音乐',
            filetypes=[('音频文件', ' '.join('*' + e for e in ALLOWED_MUSIC_EXT)),
                       ('所有文件', '*.*')])
        if not path:
            return
        src = Path(path)
        if src.stat().st_size > 25 * 1024 * 1024:
            messagebox.showwarning('背景音乐', '音频文件超过 25MB，请压缩后再试',
                                    parent=self.win)
            return
        # 复制进数据目录，只保留一份
        for old in MUSIC_DIR.glob('*'):
            if old.is_file():
                try:
                    old.unlink()
                except OSError:
                    pass
        dest = MUSIC_DIR / src.name
        try:
            import shutil
            shutil.copy2(str(src), str(dest))
        except OSError as e:
            messagebox.showerror('背景音乐', '复制失败：%s' % e, parent=self.win)
            return
        self.state['music'] = {
            'filename': dest.name,
            'original_name': src.name,
            'size': dest.stat().st_size,
            'uploaded_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        save_state(self.state)
        self._refresh_music()
        self._status('已设置背景音乐：%s' % src.name)

    def toggle_preview(self):
        path = self.get_music_path()
        if not path:
            messagebox.showinfo('背景音乐', '还没有设置背景音乐', parent=self.win)
            return
        if self._preview.playing:
            self._preview.stop()
        else:
            if not self._preview.play(path):
                messagebox.showwarning('背景音乐',
                                       '无法播放该音频文件', parent=self.win)

    def delete_music(self):
        if not self.state.get('music'):
            return
        if not messagebox.askyesno('背景音乐', '删除当前背景音乐？',
                                   parent=self.win):
            return
        self._preview.stop()
        self.state['music'] = None
        save_state(self.state)
        for old in MUSIC_DIR.glob('*'):
            if old.is_file():
                try:
                    old.unlink()
                except OSError:
                    pass
        self._refresh_music()
        self._status('已删除背景音乐')

    def _refresh_music(self):
        music = self.state.get('music') or {}
        if music.get('filename'):
            size = music.get('size') or 0
            text = '%s（%.1f MB）' % (music.get('original_name')
                                      or music['filename'],
                                      size / 1024.0 / 1024.0)
        else:
            text = '未设置'
        self.lbl_music.configure(text=text)

    # ---------- 操作区 ----------
    def _build_actions(self, parent):
        bar = tk.Frame(parent, bg=PALETTE['bg'])
        bar.pack(fill='x', pady=(0, 8))

        # 后台不提供抽签入口：抽签只能在前台大屏用空格键进行，
        # 后台只负责名单/设置/结果查看/校验/导出
        FlatButton(bar, text='打开大屏', kind='primary',
                   command=self.open_screen).pack(side='left')
        FlatButton(bar, text='重置结果', kind='danger',
                   command=self.reset_result_ui).pack(side='left', padx=8)

        right = tk.Frame(bar, bg=PALETTE['bg'])
        right.pack(side='right')
        FlatButton(right, text='导出 XLSX',
                   command=lambda: self.export_result('xlsx')).pack(side='left')
        FlatButton(right, text='导出 CSV',
                   command=lambda: self.export_result('csv')).pack(
            side='left', padx=6)
        FlatButton(right, text='复现校验', command=self.verify_current).pack(
            side='left')

        self.lbl_status = tk.Label(parent, bg=PALETTE['bg'],
                                   fg=PALETTE['muted'], font=font(9),
                                   anchor='w',
                                   text='抽签在大屏进行：打开大屏后按 空格键 开始，'
                                        'ESC 退出')
        self.lbl_status.pack(fill='x')

    def _status(self, text):
        self.lbl_status.configure(text=text)

    # ---------- 结果 ----------
    def _build_result(self, parent):
        box = ttk.Labelframe(parent, text='抽签结果', padding=6)
        box.pack(fill='both', expand=True)

        columns = ('no', 'name', 'org', 'group', 'note')
        self.tree_result = ttk.Treeview(box, columns=columns, show='headings',
                                        selectmode='browse')
        for col, text, width, anchor in (
                ('no', '出场签号', 70, 'center'),
                ('name', '姓名', 120, 'center'),
                ('org', '单位 / 部门', 170, 'w'),
                ('group', '组别', 60, 'center'),
                ('note', '备注', 150, 'w')):
            self.tree_result.heading(col, text=text)
            self.tree_result.column(col, width=width, anchor=anchor)
        scroll = ttk.Scrollbar(box, orient='vertical',
                               command=self.tree_result.yview)
        self.tree_result.configure(yscrollcommand=scroll.set)
        scroll.pack(side='right', fill='y')
        self.tree_result.pack(fill='both', expand=True)
        self.lbl_result_meta = tk.Label(box, bg=PALETTE['bg'],
                                        fg=PALETTE['faint'], font=font(9),
                                        anchor='w', text='还没有抽签结果')
        self.lbl_result_meta.pack(fill='x', pady=(4, 0))

    def refresh_result(self):
        self.tree_result.delete(*self.tree_result.get_children())
        record = self.state.get('result')
        if not record:
            self.lbl_result_meta.configure(text='还没有抽签结果')
            return
        multi = (record.get('group_count') or 0) > 1
        for item in record.get('items', []):
            group = ('第 %d 组' % item['group']
                     if multi and item.get('group') else '')
            self.tree_result.insert('', 'end', values=(
                item.get('no'), item.get('name'), item.get('org'),
                group, item.get('note')))
        self.lbl_result_meta.configure(
            text='%s · 抽签时间 %s · 种子 %s · 记录 %s'
                 % (record.get('title', ''), record.get('created_at', ''),
                    record.get('seed', ''), record.get('id', '')))

    # ---------- 历史记录 ----------
    def _build_records(self, parent):
        box = ttk.Labelframe(parent, text='历史记录', padding=6)
        box.pack(fill='x', pady=(8, 0))

        columns = ('created_at', 'count', 'title')
        self.tree_records = ttk.Treeview(box, columns=columns, show='headings',
                                         height=5, selectmode='browse')
        for col, text, width, anchor in (
                ('created_at', '时间', 150, 'center'),
                ('count', '人数', 50, 'center'),
                ('title', '标题', 260, 'w')):
            self.tree_records.heading(col, text=text)
            self.tree_records.column(col, width=width, anchor=anchor)
        self.tree_records.pack(side='left', fill='both', expand=True)
        self.tree_records.bind('<Double-1>',
                               lambda e: self.view_record())

        bar = tk.Frame(box, bg=PALETTE['bg'])
        bar.pack(side='right', fill='y')
        FlatButton(bar, text='查看', command=self.view_record).pack(
            anchor='n', pady=2)
        FlatButton(bar, text='导出',
                   command=lambda: self.export_record_selected(
                       'xlsx')).pack(anchor='n', pady=2)

    def refresh_records(self):
        self.tree_records.delete(*self.tree_records.get_children())
        for rec in list_records():
            self.tree_records.insert('', 'end', iid=rec['id'], values=(
                rec['created_at'], rec['count'], rec['title']))

    def view_record(self):
        sel = self.tree_records.selection()
        if not sel:
            return
        record = load_record(sel[0])
        if not record:
            messagebox.showwarning('历史记录', '读取记录失败', parent=self.win)
            return
        # 仅查看：不写回 state['result']
        self._viewing = record
        self.tree_result.delete(*self.tree_result.get_children())
        multi = (record.get('group_count') or 0) > 1
        for item in record.get('items', []):
            group = ('第 %d 组' % item['group']
                     if multi and item.get('group') else '')
            self.tree_result.insert('', 'end', values=(
                item.get('no'), item.get('name'), item.get('org'),
                group, item.get('note')))
        self.lbl_result_meta.configure(
            text='【历史】%s · %s · 种子 %s · 记录 %s'
                 % (record.get('title', ''), record.get('created_at', ''),
                    record.get('seed', ''), record.get('id', '')))

    def export_record_selected(self, fmt):
        sel = self.tree_records.selection()
        if not sel:
            messagebox.showinfo('导出', '请先在历史记录中选择一条', parent=self.win)
            return
        record = load_record(sel[0])
        if record:
            self._do_export(record, fmt)

    # ---------- 抽签 ----------
    def perform_draw_record(self):
        """执行抽签并落盘；失败返回 None。供管理台 / 大屏共同调用。"""
        self._flush_pending()
        parts = _clean_participants(self.state['participants'])
        if len(parts) < 2:
            return None
        groups = int(self.state['settings'].get('group_count') or 0)
        if groups and groups > len(parts):
            messagebox.showwarning('抽签', '分组数不能大于参赛人数',
                                   parent=self.win)
            return None
        seed, items = perform_draw(parts, groups)
        record = build_record(self.state, seed, items)
        self.state['result'] = record
        self._viewing = None
        save_state(self.state)
        save_record(record)
        self.refresh_result()
        self.refresh_records()
        return record

    def reset_result(self):
        self.state['result'] = None
        save_state(self.state)
        self.refresh_result()
        # 大屏正开着时同步回到待抽签（重置只允许从后台发起）
        if self.screen is not None:
            try:
                if self.screen.is_open():
                    self.screen.on_admin_reset()
            except Exception:
                pass

    def reset_result_ui(self):
        if self.state.get('result') and not messagebox.askyesno(
                '重置', '确定清除当前抽签结果？历史记录仍会保留。',
                parent=self.win):
            return
        self.reset_result()
        self._viewing = None
        self._status('已清除当前结果')

    # ---------- 大屏 ----------
    def open_screen(self):
        if self.screen and self.screen.is_open():
            self.screen._safe(ui_theme.present_window)
            try:
                self.screen.win.deiconify()
                self.screen.win.focus_force()
            except Exception:
                pass
            return self.screen
        self.screen = LotteryScreen(self.win, self)
        return self.screen

    # ---------- 导出 / 校验 ----------
    def export_result(self, fmt='xlsx'):
        record = getattr(self, '_viewing', None) or self.state.get('result')
        if not record:
            messagebox.showwarning('导出', '还没有抽签结果', parent=self.win)
            return
        self._do_export(record, fmt)

    def _do_export(self, record, fmt):
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        if fmt == 'csv':
            path = filedialog.asksaveasfilename(
                parent=self.win, defaultextension='.csv',
                initialfile='出场顺序_%s.csv' % stamp,
                filetypes=[('CSV 文件', '*.csv')])
            if not path:
                return
            try:
                with open(path, 'w', encoding='utf-8-sig', newline='') as f:
                    f.write(build_csv_text(record))
            except OSError as e:
                messagebox.showerror('导出失败', str(e), parent=self.win)
                return
        else:
            path = filedialog.asksaveasfilename(
                parent=self.win, defaultextension='.xlsx',
                initialfile='出场顺序_%s.xlsx' % stamp,
                filetypes=[('Excel 工作簿', '*.xlsx')])
            if not path:
                return
            try:
                export_xlsx(record, path)
            except ImportError:
                messagebox.showerror(
                    '导出失败', '当前环境缺少 openpyxl，无法导出 Excel。\n'
                    '可先导出 CSV。', parent=self.win)
                return
            except Exception as e:
                messagebox.showerror('导出失败', str(e), parent=self.win)
                return
        self._status('已导出：%s' % path)
        try:
            os.startfile(os.path.dirname(path))
        except Exception:
            pass

    def verify_current(self):
        record = self.state.get('result')
        if not record:
            messagebox.showinfo('复现校验', '还没有抽签结果可校验', parent=self.win)
            return
        matched = verify_record(record)
        if matched:
            messagebox.showinfo(
                '复现校验通过',
                '用种子 %s 复现出的顺序与存档完全一致。\n\n'
                '记录编号 %s · 参与 %d 人\n算法 %s'
                % (record.get('seed'), record.get('id'),
                   len(record.get('items', [])), record.get('algorithm')),
                parent=self.win)
        else:
            messagebox.showerror(
                '复现校验不通过',
                '复现结果与存档不一致（存档可能被改动）。\n记录编号 %s'
                % record.get('id'), parent=self.win)

    # ---------- 生命周期 ----------
    def refresh_all(self):
        self._last_saved_parts = json.loads(json.dumps(
            self.state.get('participants') or []))
        self.rebuild_rows()
        self.refresh_result()
        self.refresh_records()
        self._refresh_music()

    def close_requested(self):
        """用户点 X：现场抽签正在进行，先二次确认防误触。"""
        try:
            if not ui_theme.confirm_close(self.win, '现场抽签排序'):
                return
        except Exception:
            pass
        # 确认后优先交给会务主程序统一关闭（保持子窗口互斥状态一致）。
        if callable(self.on_close):
            self.on_close()
        else:
            self.close()

    def close(self):
        self._flush_pending()
        if self._save_job:
            try:
                self.win.after_cancel(self._save_job)
            except Exception:
                pass
            self._save_job = None
        self._preview.stop()
        if self.screen is not None:
            try:
                self.screen.close()
            except Exception:
                pass
            self.screen = None
        try:
            if self.win.winfo_exists():
                self.win.destroy()
        except Exception:
            pass
