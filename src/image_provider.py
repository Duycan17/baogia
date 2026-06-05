"""Product image loading with local cache."""

from __future__ import annotations

import hashlib
import io
import random
from pathlib import Path
from typing import TYPE_CHECKING, Tuple
from urllib.parse import urlparse

import requests
from PIL import Image, ImageDraw, ImageFont

if TYPE_CHECKING:
    from product_catalog import Product

_SIZE: Tuple[int, int] = (320, 200)
_ROOT = Path(__file__).resolve().parent.parent
_DEFAULT_IMAGE = _ROOT / "assets" / "default_product.jpg"
_LEGACY_IMAGE = _ROOT / "assets" / "furniture.jpg"
_CACHE_DIR = _ROOT / "data" / "product_images"


def _default_image_path() -> Path:
    if _DEFAULT_IMAGE.exists():
        return _DEFAULT_IMAGE
    return _LEGACY_IMAGE


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


def _to_png_bytes(path: Path) -> bytes:
    with Image.open(path) as img:
        normalized = img.convert("RGB")
        buf = io.BytesIO()
        normalized.save(buf, format="PNG")
        return buf.getvalue()


def _cache_path(product_code: str, suffix: str = ".jpg") -> Path:
    digest = hashlib.sha256(product_code.encode("utf-8")).hexdigest()[:16]
    return _CACHE_DIR / f"{digest}{suffix}"


def _download_image(url: str, dest: Path) -> bool:
    try:
        resp = requests.get(url, timeout=20)
        resp.raise_for_status()
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(resp.content)
        return True
    except (OSError, requests.RequestException):
        return False


def get_product_image(product: Product) -> bytes:
    ref = (product.image_ref or "").strip()
    cache = _cache_path(product.code)

    if cache.is_file():
        return _to_png_bytes(cache)

    if ref:
        parsed = urlparse(ref)
        if parsed.scheme in ("http", "https"):
            if _download_image(ref, cache):
                return _to_png_bytes(cache)
        else:
            local = Path(ref)
            if not local.is_absolute():
                local = _ROOT / ref
            if local.is_file():
                return _to_png_bytes(local)

    fallback = _default_image_path()
    if fallback.is_file():
        return _to_png_bytes(fallback)
    return random_placeholder_png()
