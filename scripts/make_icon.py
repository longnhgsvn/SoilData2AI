# -*- coding: utf-8 -*-
"""Tạo icon ứng dụng (assets/app.ico) không cần file nguồn bên ngoài."""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

HERE = Path(__file__).resolve().parent.parent
OUT = HERE / "assets"
OUT.mkdir(parents=True, exist_ok=True)


def font(size: int) -> ImageFont.FreeTypeFont:
    for p in (
        r"C:\Windows\Fonts\segoeuib.ttf",
        r"C:\Windows\Fonts\arialbd.ttf",
        r"C:\Windows\Fonts\calibrib.ttf",
    ):
        if Path(p).exists():
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


def make(size: int = 256) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    r = int(size * 0.20)
    # nền navy bo góc
    d.rounded_rectangle((2, 2, size - 2, size - 2), radius=r, fill=(11, 44, 77, 255))
    # dải xanh lá phía dưới
    d.rounded_rectangle(
        (int(size * 0.10), int(size * 0.72), int(size * 0.90), int(size * 0.88)),
        radius=int(size * 0.06),
        fill=(95, 168, 60, 255),
    )
    # mũi tên trắng
    y = int(size * 0.45)
    d.line(
        [(int(size * 0.20), y), (int(size * 0.64), y)],
        fill=(255, 255, 255, 255),
        width=max(6, int(size * 0.055)),
    )
    d.polygon(
        [
            (int(size * 0.62), int(size * 0.31)),
            (int(size * 0.86), y),
            (int(size * 0.62), int(size * 0.59)),
        ],
        fill=(255, 255, 255, 255),
    )
    # chữ CSV
    f = font(int(size * 0.28))
    text = "CSV"
    box = d.textbbox((0, 0), text, font=f)
    tw, th = box[2] - box[0], box[3] - box[1]
    d.text(
        ((size - tw) / 2, size * 0.73 - th / 2),
        text,
        font=f,
        fill=(255, 255, 255, 255),
    )
    return img


def main() -> int:
    big = make(256)
    big.save(OUT / "app.png")
    sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    big.save(OUT / "app.ico", format="ICO", sizes=sizes)
    print(f"OK: {OUT / 'app.ico'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

