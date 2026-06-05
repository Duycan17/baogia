"""Extract embedded JPEG/PNG images from legacy .xls (OLE) exports."""

from __future__ import annotations

from pathlib import Path


def _extract_jpegs(blob: bytes) -> list[tuple[int, bytes]]:
    out: list[tuple[int, bytes]] = []
    i = 0
    while True:
        start = blob.find(b"\xff\xd8\xff", i)
        if start < 0:
            break
        end = blob.find(b"\xff\xd9", start + 3)
        if end < 0:
            break
        out.append((start, blob[start : end + 2]))
        i = end + 2
    return out


def _extract_pngs(blob: bytes) -> list[tuple[int, bytes]]:
    out: list[tuple[int, bytes]] = []
    i = 0
    marker = b"\x89PNG\r\n\x1a\n"
    iend = b"IEND\xaeB`\x82"
    while True:
        start = blob.find(marker, i)
        if start < 0:
            break
        end = blob.find(iend, start)
        if end < 0:
            break
        out.append((start, blob[start : end + len(iend)]))
        i = end + len(iend)
    return out


def extract_embedded_images(path: Path) -> list[bytes]:
    """Return images sorted by byte offset in file (typical row order in AMIS export)."""
    data = path.read_bytes()
    tagged = _extract_jpegs(data) + _extract_pngs(data)
    tagged.sort(key=lambda x: x[0])
    images = [img for _, img in tagged]
    # Drop tiny false-positive JPEG markers if any (< 1KB)
    return [img for img in images if len(img) >= 1024]
