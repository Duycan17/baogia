"""Load product master data from AMIS .xls export."""

from __future__ import annotations

import hashlib
import os
import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

import xlrd

from xls_image_extractor import extract_embedded_images

_ROOT = Path(__file__).resolve().parent.parent
_DEFAULT_CATALOG = _ROOT / "file nhap quat AMIS.xls"
_IMAGE_CACHE_DIR = _ROOT / "data" / "product_images"


def _norm_header(text: str) -> str:
    return re.sub(r"\s+", " ", str(text).replace("\n", " ").strip().lower())


def _norm_code(code: str) -> str:
    return re.sub(r"\s+", "", str(code).strip()).lower()


def _safe_image_stem(code: str) -> str:
    digest = hashlib.sha256(code.encode("utf-8")).hexdigest()[:16]
    return digest


@dataclass(frozen=True)
class Product:
    code: str
    name: str
    unit: str
    sale_price: float
    description: str
    vat_percent: float
    image_ref: str


def catalog_path() -> Path:
    env = os.environ.get("CATALOG_PATH", "").strip()
    return Path(env) if env else _DEFAULT_CATALOG


def _header_map(sheet, header_row: int) -> dict[str, int]:
    """Merge headers from sub-header row (e.g. row 6 AZ) and main header row."""
    mapping: dict[str, int] = {}
    for r in range(max(0, header_row - 1), header_row + 1):
        for c in range(sheet.ncols):
            raw = str(sheet.cell_value(r, c)).strip()
            if raw:
                mapping[_norm_header(raw)] = c
    return mapping


def _col(headers: dict[str, int], *candidates: str) -> int | None:
    for name in candidates:
        key = _norm_header(name)
        if key in headers:
            return headers[key]
    return None


def _col_image(headers: dict[str, int]) -> int | None:
    for key, idx in headers.items():
        if "hình ảnh" in key and "sản phẩm" in key:
            return idx
    return _col(headers, "hình ảnh sản phẩm")


def _cell_str(sheet, row: int, col: int | None) -> str:
    if col is None:
        return ""
    val = sheet.cell_value(row, col)
    if val is None:
        return ""
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val).strip()


def _cell_float(sheet, row: int, col: int | None, default: float = 0.0) -> float:
    if col is None:
        return default
    val = sheet.cell_value(row, col)
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


@lru_cache(maxsize=2)
def load_catalog(path_str: str) -> tuple[Product, ...]:
    path = Path(path_str)
    if not path.is_file():
        raise FileNotFoundError(f"Không tìm thấy file AMIS: {path}")

    wb = xlrd.open_workbook(str(path), formatting_info=True)
    sheet = wb.sheet_by_index(0)
    header_row = next(
        (r for r in range(sheet.nrows) if "mã hàng" in _norm_header(sheet.cell_value(r, 0))),
        None,
    )
    if header_row is None:
        raise ValueError("File AMIS không có dòng tiêu đề Mã hàng")

    headers = _header_map(sheet, header_row)
    col_code = _col(headers, "mã hàng (*)", "mã hàng")
    col_name = _col(headers, "tên hàng (*)", "tên hàng")
    col_unit = _col(headers, "đơn vị tính chính", "đơn vị tính")
    col_desc = _col(headers, "mô tả")
    col_price = _col(headers, "đơn giá bán")
    col_vat = _col(headers, "thuế suất gtgt (%)", "thuế suất")
    col_image = _col_image(headers)
    embedded_images = extract_embedded_images(path)
    _IMAGE_CACHE_DIR.mkdir(parents=True, exist_ok=True)

    products: list[Product] = []
    image_idx = 0
    for r in range(header_row + 1, sheet.nrows):
        code = _cell_str(sheet, r, col_code)
        if not code:
            continue

        image_ref = _cell_str(sheet, r, col_image)
        if not image_ref and image_idx < len(embedded_images):
            blob = embedded_images[image_idx]
            ext = ".png" if blob.startswith(b"\x89PNG") else ".jpg"
            cache_file = _IMAGE_CACHE_DIR / f"{_safe_image_stem(code)}{ext}"
            if not cache_file.is_file():
                cache_file.write_bytes(blob)
            image_ref = str(cache_file.relative_to(_ROOT))
            image_idx += 1

        products.append(
            Product(
                code=code,
                name=_cell_str(sheet, r, col_name) or code,
                unit=_cell_str(sheet, r, col_unit) or "Cái",
                sale_price=_cell_float(sheet, r, col_price),
                description=_cell_str(sheet, r, col_desc),
                vat_percent=_cell_float(sheet, r, col_vat, 10.0),
                image_ref=image_ref,
            )
        )
    return tuple(products)


def _catalog() -> tuple[Product, ...]:
    return load_catalog(str(catalog_path()))


def list_codes() -> list[str]:
    return [p.code for p in _catalog()]


def lookup_by_code(code: str) -> Product | None:
    target = _norm_code(code)
    if not target:
        return None
    for product in _catalog():
        if _norm_code(product.code) == target:
            return product
    return None


def resolve_codes(codes: list[str]) -> tuple[list[Product], list[str]]:
    found: list[Product] = []
    missing: list[str] = []
    for code in codes:
        product = lookup_by_code(code)
        if product is None:
            missing.append(code)
        else:
            found.append(product)
    return found, missing
