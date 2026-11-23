"""生成安装程序图标（纯 Python，不依赖 PIL）：Office 蓝圆角方块 + 三条版本层。

用法：python make-icon.py app.ico
"""
import struct
import sys


def blend(a, b, t):
    return tuple(int(round(a[i] + (b[i] - a[i]) * t)) for i in range(3))


def rounded_rect_alpha(x, y, left, top, right, bottom, radius):
    """点 (x, y) 在圆角矩形内的覆盖率（0/1，采样点已做过超采样）。"""
    if x < left or x >= right or y < top or y >= bottom:
        return 0.0
    dx = 0.0
    dy = 0.0
    if x < left + radius:
        dx = left + radius - x
    elif x > right - radius:
        dx = x - (right - radius)
    if y < top + radius:
        dy = top + radius - y
    elif y > bottom - radius:
        dy = y - (bottom - radius)
    if dx > 0 and dy > 0:
        return 1.0 if dx * dx + dy * dy <= radius * radius else 0.0
    return 1.0


def render(size):
    scale = size / 256.0
    # 超采样 3x3 抗锯齿
    ss = 3
    pixels = []
    for py in range(size):
        row = []
        for px in range(size):
            r = g = b = a = 0.0
            for sy in range(ss):
                for sx in range(ss):
                    x = (px + (sx + 0.5) / ss) / scale
                    y = (py + (sy + 0.5) / ss) / scale
                    if rounded_rect_alpha(x, y, 2, 2, 254, 254, 46) > 0:
                        # 背景：上亮下深的 Office 蓝
                        base = blend((31, 111, 209), (10, 74, 148), min(1.0, max(0.0, (y - 2) / 252.0)))
                        layer = (base, 1.0)
                    else:
                        layer = ((0, 0, 0), 0.0)
                    # 三条白色横条（代表版本层）
                    for top, width, height in ((66, 150, 26), (112, 150, 26), (158, 104, 26)):
                        left = (256 - width) / 2.0
                        if rounded_rect_alpha(x, y, left, top, left + width, top + height, height / 2.0) > 0:
                            layer = ((255, 255, 255), 0.93)
                            break
                    r += layer[0][0] * layer[1]
                    g += layer[0][1] * layer[1]
                    b += layer[0][2] * layer[1]
                    a += layer[1]
            count = ss * ss
            row.append((int(r / count), int(g / count), int(b / count), int(255 * a / count)))
        pixels.append(row)
    return pixels


def write_ico(path, sizes):
    images = []
    for size in sizes:
        pixels = render(size)
        header = struct.pack("<IiiHHIIiiII", 40, size, size * 2, 1, 32, 0, 0, 0, 0, 0, 0)
        body = bytearray()
        for y in range(size - 1, -1, -1):
            for x in range(size):
                r, g, b, a = pixels[y][x]
                body += bytes((b, g, r, a))
        # AND 掩码：32bpp 用 alpha 通道，掩码全 0，行按 4 字节对齐
        row_bytes = ((size + 31) // 32) * 4
        mask = bytes(row_bytes * size)
        images.append(header + bytes(body) + mask)

    header = struct.pack("<HHH", 0, 1, len(sizes))
    offset = 6 + 16 * len(sizes)
    entries = b""
    for size, data in zip(sizes, images):
        entries += struct.pack("<BBBBHHII", 0 if size >= 256 else size, 0 if size >= 256 else size,
                               0, 0, 1, 32, len(data), offset)
        offset += len(data)
    with open(path, "wb") as file:
        file.write(header + entries + b"".join(images))
    print("wrote", path, "sizes", sizes)


if __name__ == "__main__":
    write_ico(sys.argv[1] if len(sys.argv) > 1 else "app.ico", [16, 24, 32, 48, 64, 128, 256])
