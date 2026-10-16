"""生成 README 用的案例图：上排是模板原页，下排是生成结果的精选页。
用法：python3 tools/gallery.py finance 1,4,5,7,11,13
"""
import glob
import os
import re
import sys

from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def font(size):
    for p in ("/System/Library/Fonts/PingFang.ttc", "/System/Library/Fonts/Hiragino Sans GB.ttc",
              "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"):
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def pages(d):
    fs = glob.glob(os.path.join(d, "s-*.jpg"))
    return sorted(fs, key=lambda f: int(re.findall(r"(\d+)\.jpg$", f)[0]))


def main(ex, picks):
    tpl = pages(os.path.join(ROOT, "examples", ex, "work", "template-render"))
    out = pages(os.path.join(ROOT, "examples", ex, "out", "render"))
    W, gap, pad = 1800, 14, 36
    cols = 3
    ow = (W - 2 * pad - gap * (cols - 1)) // cols
    oh = ow * 9 // 16
    tw = (W - 2 * pad - gap * (len(tpl) - 1)) // max(len(tpl), 6)
    th = tw * 9 // 16
    rows = (len(picks) + cols - 1) // cols
    H = pad + 44 + th + pad + 44 + rows * oh + (rows - 1) * gap + pad
    im = Image.new("RGB", (W, H), (245, 245, 247))
    d = ImageDraw.Draw(im)
    f = font(26)
    d.text((pad, pad), f"模板原页（{len(tpl)} 页）", fill=(90, 90, 96), font=f)
    y = pad + 44
    for i, p in enumerate(tpl):
        im.paste(Image.open(p).convert("RGB").resize((tw, th), Image.LANCZOS), (pad + i * (tw + gap), y))
    y += th + pad
    d.text((pad, y), f"生成结果（共 {len(out)} 页，节选）", fill=(90, 90, 96), font=f)
    y += 44
    for k, n in enumerate(picks):
        r, c = divmod(k, cols)
        im.paste(Image.open(out[n - 1]).convert("RGB").resize((ow, oh), Image.LANCZOS),
                 (pad + c * (ow + gap), y + r * (oh + gap)))
    dst = os.path.join(ROOT, "docs", "gallery", f"{ex}.jpg")
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    im.save(dst, quality=84, optimize=True)
    print(dst, im.size)


if __name__ == "__main__":
    main(sys.argv[1], [int(x) for x in sys.argv[2].split(",")])
