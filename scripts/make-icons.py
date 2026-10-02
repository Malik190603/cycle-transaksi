# Membuat ikon peluncur dan splash (dijalankan manual bila ikon diganti): python3 scripts/make-icons.py
from PIL import Image, ImageDraw
import os
RES = os.path.join(os.path.dirname(__file__), '..', 'android', 'app', 'src', 'main', 'res')
BG = (10, 108, 255, 255); BG2 = (0, 82, 204, 255); WHITE = (255, 255, 255, 255)
S = 1024

def glyph(scale=1.0):
    """Papan klip putih dengan tanda centang, di kanvas transparan S x S."""
    im = Image.new('RGBA', (S, S), (0, 0, 0, 0)); d = ImageDraw.Draw(im)
    c = S / 2; w = 400 * scale; h = 500 * scale
    x0, y0, x1, y1 = c - w / 2, c - h / 2 + 24 * scale, c + w / 2, c + h / 2 + 24 * scale
    d.rounded_rectangle((x0, y0, x1, y1), radius=56 * scale, fill=WHITE)
    cw, ch = 190 * scale, 84 * scale  # penjepit
    d.rounded_rectangle((c - cw / 2, y0 - ch / 2, c + cw / 2, y0 + ch / 2), radius=30 * scale, fill=WHITE, outline=BG, width=int(18 * scale))
    lw = int(54 * scale)  # centang
    pts = [(c - 110 * scale, c + 50 * scale), (c - 30 * scale, c + 130 * scale), (c + 120 * scale, c - 50 * scale)]
    d.line(pts, fill=BG, width=lw, joint='curve')
    for p in pts: d.ellipse((p[0] - lw / 2, p[1] - lw / 2, p[0] + lw / 2, p[1] + lw / 2), fill=BG)
    return im

def gradient(size):
    im = Image.new('RGBA', (1, 256))
    for y in range(256):
        t = y / 255
        im.putpixel((0, y), tuple(int(BG[i] + (BG2[i] - BG[i]) * t) for i in range(3)) + (255,))
    return im.resize((size, size))

def save(im, path, size):
    im.resize(size if isinstance(size, tuple) else (size, size), Image.LANCZOS).save(os.path.join(RES, path))

fg = glyph(0.92)                       # lapis depan ikon adaptif (aman di zona 66%)
legacy = gradient(S); legacy.alpha_composite(glyph(1.25))
mask = Image.new('L', (S, S), 0); ImageDraw.Draw(mask).rounded_rectangle((0, 0, S, S), radius=220, fill=255)
sq = Image.new('RGBA', (S, S), (0, 0, 0, 0)); sq.paste(legacy, (0, 0), mask)
cm = Image.new('L', (S, S), 0); ImageDraw.Draw(cm).ellipse((0, 0, S, S), fill=255)
rd = Image.new('RGBA', (S, S), (0, 0, 0, 0)); rd.paste(legacy, (0, 0), cm)
for name, px in [('mdpi', 48), ('hdpi', 72), ('xhdpi', 96), ('xxhdpi', 144), ('xxxhdpi', 192)]:
    save(sq, f'mipmap-{name}/ic_launcher.png', px); save(rd, f'mipmap-{name}/ic_launcher_round.png', px)
    save(fg, f'mipmap-{name}/ic_launcher_foreground.png', int(px * 2.25))

def splash(w, h):
    im = Image.new('RGBA', (w, h), BG); g = int(min(w, h) * 0.42)
    im.alpha_composite(glyph(1.25).resize((g, g), Image.LANCZOS), ((w - g) // 2, (h - g) // 2)); return im.convert('RGB')
for name, (a, b) in {'mdpi': (320, 480), 'hdpi': (480, 800), 'xhdpi': (720, 1280), 'xxhdpi': (960, 1600), 'xxxhdpi': (1280, 1920)}.items():
    splash(a, b).save(os.path.join(RES, f'drawable-port-{name}/splash.png')); splash(b, a).save(os.path.join(RES, f'drawable-land-{name}/splash.png'))
splash(480, 320).save(os.path.join(RES, 'drawable/splash.png'))
sq.resize((512, 512), Image.LANCZOS).save(os.path.join(os.path.dirname(__file__), '..', 'icon-preview.png'))
print('ikon & splash dibuat')
