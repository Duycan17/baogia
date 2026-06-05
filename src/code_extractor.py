"""Extract product codes referenced in an uploaded báo giá .docx."""

from __future__ import annotations

import re

from docx import Document

from docx_utils import iter_paragraph_texts
from product_catalog import list_codes

_EXPLICIT_RE = re.compile(
    r"(?:Mã|Ma)\s*hàng\s*[:：]\s*(.+)",
    re.IGNORECASE | re.UNICODE,
)


def _split_codes(blob: str) -> list[str]:
    # Semicolon/newline only — AMIS codes may contain commas.
    parts = re.split(r"[;\n]+", blob)
    return [p.strip() for p in parts if p.strip()]


def _scan_catalog_tokens(text: str, catalog_codes: list[str]) -> list[str]:
    """Longest catalog code matches first to avoid partial overlaps."""
    ordered = sorted(catalog_codes, key=len, reverse=True)
    lower_text = text.lower()
    found: list[str] = []
    used_spans: list[tuple[int, int]] = []

    for code in ordered:
        start = 0
        needle = code.lower()
        while True:
            idx = lower_text.find(needle, start)
            if idx < 0:
                break
            end = idx + len(needle)
            if not any(a < end and b > idx for a, b in used_spans):
                found.append(code)
                used_spans.append((idx, end))
            start = idx + 1
    return found


def extract_product_codes(doc: Document) -> list[str]:
    catalog_codes = list_codes()
    catalog_set = {c.lower() for c in catalog_codes}
    explicit: list[str] = []
    full_text_parts: list[str] = []

    for text in iter_paragraph_texts(doc):
        full_text_parts.append(text)
        m = _EXPLICIT_RE.search(text)
        if m:
            explicit.extend(_split_codes(m.group(1)))

    deduped_explicit: list[str] = []
    seen: set[str] = set()
    for code in explicit:
        key = code.lower()
        if key not in seen:
            seen.add(key)
            deduped_explicit.append(code)

    if deduped_explicit:
        return deduped_explicit

    scanned = _scan_catalog_tokens("\n".join(full_text_parts), catalog_codes)
    result: list[str] = []
    seen2: set[str] = set()
    for code in scanned:
        if code.lower() in catalog_set and code.lower() not in seen2:
            seen2.add(code.lower())
            result.append(code)
    return result
