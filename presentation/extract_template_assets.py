# -*- coding: utf-8 -*-
"""Trích nền và logo từ file PPTX mẫu ra thư mục assets/."""
from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "templates" / "Side 60 năm_ERSD 3.pptx"
ASSETS = HERE / "assets"


def main() -> int:
    if not TEMPLATE.exists():
        raise FileNotFoundError(TEMPLATE)
    ASSETS.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(TEMPLATE) as z:
        # image1.png: nền full-slide (watermark + logo 60 năm)
        # image2.png: logo ERSD góc trên trái
        shutil.copyfileobj(z.open("ppt/media/image1.png"), (ASSETS / "background.png").open("wb"))
        shutil.copyfileobj(z.open("ppt/media/image2.png"), (ASSETS / "ersd_logo.png").open("wb"))
    print(f"OK: {ASSETS / 'background.png'}")
    print(f"OK: {ASSETS / 'ersd_logo.png'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

