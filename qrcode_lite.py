# -*- coding: utf-8 -*-
"""极简二维码生成（纯 Python，不依赖第三方库）。

用于两种用途：
  * “手机扫码打开局域网地址”（字节模式，纠错 M）
  * 二维码生成器：链接 / 普通文本，可加中央 logo（纠错 H，
    放不下时自动降级 M → L）
  * 版本 1-6（H 级支持 1-5，最长约 44 字节；M/L 最长约 106/134 字节）

用法：
    matrix = make_matrix('http://192.168.1.20:8000')
    img = to_image(matrix, scale=8)          # PIL.Image（黑白）
    img = paste_logo(img, 'logo.png')        # 白色圆角底 + 中央 logo

H 级纠错表说明：v1-5 的 (块数, 每块数据码字) 与 ISO 18004 一致，
并用 OpenCV QRCodeDetector 解码回环验证过（错误结构均无法解码）；
v6 的 H 表未验证，故不放，长内容自动降级 M。
"""

_EXP = [0] * 512
_LOG = [0] * 256


def _init_tables():
    x = 1
    for i in range(255):
        _EXP[i] = x
        _LOG[x] = i
        x <<= 1
        if x & 0x100:
            x ^= 0x11D
    for i in range(255, 512):
        _EXP[i] = _EXP[i - 255]


_init_tables()


def _mul(a, b):
    if a == 0 or b == 0:
        return 0
    return _EXP[_LOG[a] + _LOG[b]]


def _rs_generator(n):
    """Reed-Solomon 生成多项式（系数从高次到低次）。"""
    poly = [1]
    for i in range(n):
        new = [0] * (len(poly) + 1)
        for j, coef in enumerate(poly):
            new[j] ^= coef
            new[j + 1] ^= _mul(coef, _EXP[i])
        poly = new
    return poly


def _rs_encode(data, ec_len):
    gen = _rs_generator(ec_len)
    msg = list(data) + [0] * ec_len
    for i in range(len(data)):
        coef = msg[i]
        if coef:
            for j, g in enumerate(gen):
                msg[i + j] ^= _mul(g, coef)
    return msg[-ec_len:]


# (纠错等级, 版本) -> (每块纠错码字数, [(块数, 每块数据码字数), ...])
_EC_TABLE = {
    ('M', 1): (10, [(1, 16)]),
    ('M', 2): (16, [(1, 28)]),
    ('M', 3): (26, [(1, 44)]),
    ('M', 4): (18, [(2, 32)]),
    ('M', 5): (24, [(2, 43)]),
    ('M', 6): (16, [(4, 27)]),
    ('L', 1): (7, [(1, 19)]),
    ('L', 2): (10, [(1, 34)]),
    ('L', 3): (15, [(1, 55)]),
    ('L', 4): (20, [(1, 80)]),
    ('L', 5): (26, [(1, 108)]),
    ('L', 6): (18, [(2, 68)]),
    # H 级：加 logo 时遮住中央区域，需要更强的纠错（可恢复 ~30%）。
    # 只放验证过的 v1-5；v6 未验证，放不下时 _pick_version 会降级 M。
    ('H', 1): (17, [(1, 9)]),
    ('H', 2): (28, [(1, 16)]),
    ('H', 3): (22, [(2, 13)]),
    ('H', 4): (16, [(4, 9)]),
    ('H', 5): (22, [(2, 11), (2, 12)]),
}

_ALIGN = {1: [], 2: [6, 18], 3: [6, 22], 4: [6, 26], 5: [6, 30], 6: [6, 34]}

_ECC_BITS = {'L': 0b01, 'M': 0b00, 'Q': 0b11, 'H': 0b10}


def _data_capacity(ecc, version):
    _, blocks = _EC_TABLE[(ecc, version)]
    return sum(count * size for count, size in blocks)


def _pick_version(text, ecc='M'):
    raw = len(text.encode('utf-8'))
    # 降级链：H → M → L（Q 级没放表）。低纠错放得下长内容。
    chain = [ecc] + [lv for lv in ('M', 'L') if lv != ecc]
    for ecc_level in chain:
        for version in range(1, 7):
            # 4 位模式 + 8 位长度 + 数据，再留 4 位终止符
            if (ecc_level, version) not in _EC_TABLE:
                continue
            if raw + 2 <= _data_capacity(ecc_level, version):
                return ecc_level, version
    raise ValueError('内容太长，放不进二维码（最多约 134 字节）')


def _encode_data(text, capacity):
    data = text.encode('utf-8')
    bits = []

    def add(value, length):
        for i in range(length - 1, -1, -1):
            bits.append((value >> i) & 1)

    add(0b0100, 4)            # 字节模式
    add(len(data), 8)         # 版本 1-9 用 8 位长度
    for byte in data:
        add(byte, 8)
    total = capacity * 8
    add(0, max(0, min(4, total - len(bits))))     # 终止符
    while len(bits) % 8:
        bits.append(0)                            # 补齐到字节
    pad = (0xEC, 0x11)
    i = 0
    while len(bits) < total:
        add(pad[i % 2], 8)
        i += 1
    out = bytearray()
    for i in range(0, len(bits), 8):
        byte = 0
        for bit in bits[i:i + 8]:
            byte = (byte << 1) | bit
        out.append(byte)
    return out


def _interleave(codewords, ec_len, blocks):
    data_blocks = []
    ec_blocks = []
    pos = 0
    for count, size in blocks:
        for _ in range(count):
            block = list(codewords[pos:pos + size])
            pos += size
            data_blocks.append(block)
            ec_blocks.append(_rs_encode(block, ec_len))
    out = []
    longest = max(len(b) for b in data_blocks)
    for i in range(longest):
        for block in data_blocks:
            if i < len(block):
                out.append(block[i])
    for i in range(ec_len):
        for block in ec_blocks:
            out.append(block[i])
    return out


def _mask_bit(mask, row, col):
    if mask == 0:
        return (row + col) % 2 == 0
    if mask == 1:
        return row % 2 == 0
    if mask == 2:
        return col % 3 == 0
    if mask == 3:
        return (row + col) % 3 == 0
    if mask == 4:
        return (row // 2 + col // 3) % 2 == 0
    if mask == 5:
        return (row * col) % 2 + (row * col) % 3 == 0
    if mask == 6:
        return ((row * col) % 2 + (row * col) % 3) % 2 == 0
    return ((row + col) % 2 + (row * col) % 3) % 2 == 0


def _format_bits(ecc, mask):
    data = (_ECC_BITS[ecc] << 3) | mask
    rem = data
    for _ in range(10):
        rem = (rem << 1) ^ ((rem >> 9) * 0x537)
    value = ((data << 10) | rem) ^ 0x5412
    return [(value >> i) & 1 for i in range(14, -1, -1)]


def _build_matrix(ecc, version, codewords, mask):
    size = 17 + 4 * version
    matrix = [[0] * size for _ in range(size)]
    reserved = [[0] * size for _ in range(size)]

    def put(row, col, value):
        if 0 <= row < size and 0 <= col < size:
            matrix[row][col] = value
            reserved[row][col] = 1

    def finder(row, col):
        for r in range(-1, 8):
            for c in range(-1, 8):
                inside = 0 <= r <= 6 and 0 <= c <= 6
                edge = inside and (r in (0, 6) or c in (0, 6))
                core = 2 <= r <= 4 and 2 <= c <= 4
                put(row + r, col + c, 1 if (edge or core) and inside else 0)

    finder(0, 0)
    finder(size - 7, 0)
    finder(0, size - 7)

    # 定位图形（黑白相间）
    for i in range(8, size - 8):
        bit = 1 if i % 2 == 0 else 0
        put(6, i, bit)
        put(i, 6, bit)

    # 校正图形
    centers = _ALIGN[version]
    for r in centers:
        for c in centers:
            if (r, c) in ((6, 6), (6, size - 7), (size - 7, 6)):
                continue
            for dr in range(-2, 3):
                for dc in range(-2, 3):
                    edge = max(abs(dr), abs(dc)) == 2
                    core = dr == 0 and dc == 0
                    put(r + dr, c + dc, 1 if (edge or core) else 0)

    # 格式信息的两份位置（j 为规范里的位序，从低位开始）
    fmt_cells = []
    for i in range(15):
        j = 14 - i
        if j < 6:
            row = j
        elif j < 8:
            row = j + 1
        else:
            row = size - 15 + j
        if j < 8:
            col = size - j - 1
        elif j == 8:
            col = 7
        else:
            col = 15 - j - 1
        fmt_cells.append((i, (row, 8), (8, col)))
    # 预留格式信息区（注意不要动定时图形的 (6,8) / (8,6)）+ 固定黑点
    for _i, cell_a, cell_b in fmt_cells:
        put(cell_a[0], cell_a[1], 0)
        put(cell_b[0], cell_b[1], 0)
    put(size - 8, 8, 1)

    # 按之字形填数据
    bits = []
    for byte in codewords:
        for i in range(7, -1, -1):
            bits.append((byte >> i) & 1)
    index = 0
    col = size - 1
    upward = True
    while col > 0:
        if col == 6:
            col -= 1
        rows = range(size - 1, -1, -1) if upward else range(size)
        for row in rows:
            for c in (col, col - 1):
                if not reserved[row][c]:
                    bit = bits[index] if index < len(bits) else 0
                    index += 1
                    matrix[row][c] = 1 if (
                        bit ^ (1 if _mask_bit(mask, row, c) else 0)) else 0
        col -= 2
        upward = not upward

    # 写入格式信息（两份，位置和上面预留的一致）
    fmt = _format_bits(ecc, mask)
    for i, cell_a, cell_b in fmt_cells:
        matrix[cell_a[0]][cell_a[1]] = fmt[i]
        matrix[cell_b[0]][cell_b[1]] = fmt[i]
    return matrix


def _penalty(matrix):
    """二维码规范里的 4 条评分规则，用来选最不容易看错的掩码。"""
    size = len(matrix)
    score = 0
    # 规则 1：同色连续 5 个以上
    for line in list(matrix) + [list(col) for col in zip(*matrix)]:
        run = 1
        for i in range(1, size):
            if line[i] == line[i - 1]:
                run += 1
            else:
                if run >= 5:
                    score += run - 2
                run = 1
        if run >= 5:
            score += run - 2
    # 规则 2：2x2 同色方块
    for r in range(size - 1):
        for c in range(size - 1):
            v = matrix[r][c]
            if v == matrix[r][c + 1] == matrix[r + 1][c] == matrix[r + 1][c + 1]:
                score += 3
    # 规则 3：类似定位图形的黑白相间
    pattern = [1, 0, 1, 1, 1, 0, 1, 0, 0, 0, 0]
    for line in list(matrix) + [list(col) for col in zip(*matrix)]:
        for i in range(size - 11):
            chunk = line[i:i + 11]
            if chunk == pattern or chunk == pattern[::-1]:
                score += 40
    # 规则 4：黑白比例偏离
    dark = sum(sum(row) for row in matrix)
    ratio = dark * 100 / (size * size)
    score += 10 * (int(abs(ratio - 50) / 5))
    return score


def make_matrix(text, ecc='M'):
    """把文本编码成二维码矩阵：返回 [[0/1, ...], ...]。"""
    if not text:
        raise ValueError('内容不能为空')
    level, version = _pick_version(text, ecc)
    ec_len, blocks = _EC_TABLE[(level, version)]
    capacity = _data_capacity(level, version)
    codewords = _interleave(_encode_data(text, capacity), ec_len, blocks)
    best = None
    best_score = None
    for mask in range(8):
        matrix = _build_matrix(level, version, codewords, mask)
        score = _penalty(matrix)
        if best_score is None or score < best_score:
            best, best_score = matrix, score
    return best


def to_image(matrix, scale=6, border=6):
    """渲染成 PIL 图片（黑白，最近邻放大，扫码更清晰）。

    border 是四周的空白（静默区），规范要求至少 4 个模块；
    实测留 6 个模块时，各种长度的二维码都能被稳定识别。
    """
    from PIL import Image
    size = len(matrix)
    side = (size + border * 2) * scale
    img = Image.new('L', (side, side), 255)
    pixels = img.load()
    for r, row in enumerate(matrix):
        for c, value in enumerate(row):
            if not value:
                continue
            x0 = (c + border) * scale
            y0 = (r + border) * scale
            for y in range(y0, y0 + scale):
                for x in range(x0, x0 + scale):
                    pixels[x, y] = 0
    return img


def paste_logo(img, logo_path, ratio=0.22):
    """在二维码图片中央贴 logo：白色圆角底 + 居中 logo。

    ratio 是白底相对图片边长的比例（0.22 ≈ 遮住 5% 面积，
    配合 H 级纠错绰绰有余）。返回 RGB 图片（可直接 .save()）。
    """
    from PIL import Image, ImageDraw
    img = img.convert('RGB')
    w, h = img.size
    patch = max(int(min(w, h) * ratio) // 2 * 2, 16)
    # 白色圆角底（比 logo 大一点，隔开黑白模块）
    patch_img = Image.new('RGBA', (patch, patch), (0, 0, 0, 0))
    draw = ImageDraw.Draw(patch_img)
    draw.rounded_rectangle([0, 0, patch - 1, patch - 1],
                           radius=max(2, patch // 6),
                           fill=(255, 255, 255, 255))
    logo = Image.open(logo_path)
    if logo.mode != 'RGBA':
        logo = logo.convert('RGBA')
    # 等比缩放到白底的 78%，保持 logo 比例
    logo.thumbnail((int(patch * 0.78), int(patch * 0.78)))
    patch_img.alpha_composite(logo, ((patch - logo.size[0]) // 2,
                                    (patch - logo.size[1]) // 2))
    img.paste(patch_img, ((w - patch) // 2, (h - patch) // 2), patch_img)
    return img
