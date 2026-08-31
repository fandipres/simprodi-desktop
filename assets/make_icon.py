"""
Skrip sekali-pakai: bikin assets/icon.ico (multi-resolusi) untuk SIMPRODI
Desktop. Desain: rounded square gradient biru + monogram "SD" putih (dari
SIMPRODI Desktop) - flat & modern, tetap jelas terbaca di ukuran kecil
(16x16 taskbar).

Jalankan: python assets/make_icon.py
"""

import os

from PIL import Image, ImageDraw, ImageFont

SIZE = 256
BLUE_TOP = (37, 99, 235)      # #2563eb
BLUE_BOTTOM = (29, 78, 216)   # #1d4ed8
WHITE = (255, 255, 255, 255)


def _rounded_mask(size, radius):
    mask = Image.new("L", (size, size), 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle((0, 0, size - 1, size - 1), radius=radius, fill=255)
    return mask


def _find_font(candidates, size):
    for name in candidates:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def build_base_icon():
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))

    # Gradient vertikal biru
    gradient = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    for y in range(SIZE):
        t = y / (SIZE - 1)
        r = int(BLUE_TOP[0] + (BLUE_BOTTOM[0] - BLUE_TOP[0]) * t)
        g = int(BLUE_TOP[1] + (BLUE_BOTTOM[1] - BLUE_TOP[1]) * t)
        b = int(BLUE_TOP[2] + (BLUE_BOTTOM[2] - BLUE_TOP[2]) * t)
        ImageDraw.Draw(gradient).line([(0, y), (SIZE, y)], fill=(r, g, b, 255))

    mask = _rounded_mask(SIZE, radius=int(SIZE * 0.22))
    img.paste(gradient, (0, 0), mask)

    draw = ImageDraw.Draw(img)

    # Monogram "SD" (Sistem Dosen) di tengah, bold, putih
    font = _find_font(
        [r"C:\Windows\Fonts\segoeuib.ttf", r"C:\Windows\Fonts\arialbd.ttf"],
        int(SIZE * 0.42),
    )
    text = "SD"
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    pos = ((SIZE - tw) / 2 - bbox[0], (SIZE - th) / 2 - bbox[1])
    draw.text(pos, text, font=font, fill=WHITE)

    # Garis aksen tipis di bawah monogram (kesan "dokumen/garis SP")
    line_y = int(SIZE * 0.72)
    line_w = int(SIZE * 0.30)
    draw.rounded_rectangle(
        (SIZE / 2 - line_w / 2, line_y, SIZE / 2 + line_w / 2, line_y + int(SIZE * 0.035)),
        radius=int(SIZE * 0.02),
        fill=(255, 255, 255, 200),
    )

    return img


def main():
    out_dir = os.path.dirname(os.path.abspath(__file__))
    base = build_base_icon()
    base.save(os.path.join(out_dir, "icon_preview.png"))
    sizes = [16, 24, 32, 48, 64, 128, 256]
    base.save(
        os.path.join(out_dir, "icon.ico"),
        sizes=[(s, s) for s in sizes],
    )
    print("Saved icon.ico +", "sizes", sizes)


if __name__ == "__main__":
    main()
