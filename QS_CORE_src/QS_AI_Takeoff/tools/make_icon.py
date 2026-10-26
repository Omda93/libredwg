"""Writes config/qs_core.ico - the QS CORE mark (lid + two stacked layers, orange on dark navy) used by the
desktop / start-menu shortcuts. Pure Python (zlib PNG frames inside an ICO), no Pillow needed."""
import os
import struct
import zlib

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(os.path.dirname(HERE), "config", "qs_core.ico")
BG, FG = (0x12, 0x18, 0x26), (0xF5, 0x9E, 0x0B)


def rgba(size):
    s = size * 0.64
    x0 = (size - s) / 2
    t, g = s * 0.19, s * 0.085
    lid = t * 1.25
    y0 = (size - (lid + 2 * g + 2 * t)) / 2
    rad = size * 0.22

    def in_bg(px, py):
        cx = min(max(px, rad), size - rad)
        cy = min(max(py, rad), size - rad)
        return (px - cx) ** 2 + (py - cy) ** 2 <= rad * rad

    def in_mark(px, py):
        if y0 <= py <= y0 + lid:
            inset = s * .17 * (1 - (py - y0) / lid)
            return x0 + inset <= px <= x0 + s - inset
        for yy in (y0 + lid + g, y0 + lid + g + t + g):
            if yy <= py <= yy + t:
                r = t / 2
                cx = min(max(px, x0 + r), x0 + s - r)
                return (px - cx) ** 2 + (py - (yy + r)) ** 2 <= r * r
        return False
    ss = [(a + .5) / 4 for a in range(4)]
    rows = []
    for j in range(size):
        row = bytearray([0])
        for i in range(size):
            nb = nm = 0
            for dx in ss:
                for dy in ss:
                    if in_bg(i + dx, j + dy):
                        nb += 1
                        nm += in_mark(i + dx, j + dy)
            if not nb:
                row += bytes(4)
                continue
            k = nm / nb
            row += bytes(int(BG[c] + (FG[c] - BG[c]) * k) for c in range(3)) + bytes([int(255 * nb / 16)])
        rows.append(bytes(row))
    return b"".join(rows)


def png(size):
    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)) +
            chunk(b"IDAT", zlib.compress(rgba(size), 9)) + chunk(b"IEND", b""))


def main():
    sizes = [256, 64, 48, 32, 16]
    frames = [png(n) for n in sizes]
    head = struct.pack("<HHH", 0, 1, len(frames))
    off = 6 + 16 * len(frames)
    dirs = b""
    for n, f in zip(sizes, frames):
        dirs += struct.pack("<BBBBHHII", n % 256, n % 256, 0, 0, 1, 32, len(f), off)
        off += len(f)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "wb") as fh:
        fh.write(head + dirs + b"".join(frames))
    print(OUT)


if __name__ == "__main__":
    main()
