"""Product image providers for report embedding."""

from __future__ import annotations

import io
import random
from pathlib import Path
from typing import Tuple

from PIL import Image, ImageDraw, ImageFont

_SIZE: Tuple[int, int] = (320, 200)
_FIXED_IMAGE_PATH = Path(__file__).resolve().parent.parent / "assets" / "furniture.jpg"


def random_placeholder_png(size: Tuple[int, int] = _SIZE) -> bytes:
    w, h = size
    color = (
        random.randint(40, 220),
        random.randint(40, 220),
        random.randint(40, 220),
    )
    img = Image.new("RGB", (w, h), color)
    draw = ImageDraw.Draw(img)
    text = "Sample"
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", 28)
    except OSError:
        font = ImageFont.load_default()
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    draw.text(((w - tw) / 2, (h - th) / 2), text, fill=(20, 20, 20), font=font)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def fixed_product_image_png() -> bytes:
    """Load and normalize the fixed furniture image as PNG bytes."""
    if _FIXED_IMAGE_PATH.exists():
        with Image.open(_FIXED_IMAGE_PATH) as img:
            normalized = img.convert("RGB")
            buf = io.BytesIO()
            normalized.save(buf, format="PNG")
            return buf.getvalue()
    return random_placeholder_png()
