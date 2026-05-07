"""Load Báo Giá .docx, inject images, append discount note."""

from __future__ import annotations

import io
import re
from typing import Iterable

from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Inches
from docx.table import Table
from docx.text.paragraph import Paragraph

from image_provider import fixed_product_image_png

_CUSTOMER_ID_RE = re.compile(
    r"(?:Mã|Ma)\s*KH\s*[:：]\s*(\S+)",
    re.IGNORECASE | re.UNICODE,
)


def _iter_paragraph_texts(doc: Document) -> Iterable[str]:
    for p in doc.paragraphs:
        yield p.text
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    yield para.text


def extract_customer_id(doc: Document) -> str | None:
    for text in _iter_paragraph_texts(doc):
        m = _CUSTOMER_ID_RE.search(text)
        if m:
            return m.group(1).strip()
    return None


def extract_customer_id_from_bytes(data: bytes) -> str | None:
    return extract_customer_id(Document(io.BytesIO(data)))


def _find_items_table(doc: Document) -> Table | None:
    for table in doc.tables:
        if not table.rows:
            continue
        row0 = table.rows[0]
        joined = " ".join(c.text for c in row0.cells).lower()
        if "tt" in joined and "hình ảnh" in joined:
            return table
        if row0.cells[0].text.strip().upper().startswith("TT") and any(
            "hình" in c.text.lower() for c in row0.cells
        ):
            return table
    return None


def _image_column_index(table: Table) -> int:
    header = table.rows[0]
    for i, cell in enumerate(header.cells):
        if "hình" in cell.text.lower() or "ghi chú" in cell.text.lower():
            return i
    return len(header.cells) - 1


def _is_summary_row(row) -> bool:
    first = row.cells[0].text.strip().lower()
    return first.startswith("cộng tiền") or first.startswith("tổng tiền")


def _clear_cell_for_picture(cell) -> None:
    # Keep a single empty paragraph for the new run
    for extra in cell.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)
    p = cell.paragraphs[0]
    for el in list(p._p):
        p._p.remove(el)


def apply_image_to_items(doc: Document) -> int:
    table = _find_items_table(doc)
    if table is None:
        return 0
    col = _image_column_index(table)
    png = fixed_product_image_png()
    count = 0
    for r in range(1, len(table.rows)):
        row = table.rows[r]
        if _is_summary_row(row):
            break
        cell = row.cells[col]
        _clear_cell_for_picture(cell)
        run = cell.paragraphs[0].add_run()
        run.add_picture(io.BytesIO(png), width=Inches(1.2))
        count += 1
    return count


def append_discount_note(doc: Document, customer_id: str, discount_percent: float) -> None:
    table = _find_items_table(doc)
    customer_text = f"Mã KH: {customer_id}"
    discount_text = f"Chiết khấu áp dụng: {discount_percent:.1f}%"
    if table is None:
        p1 = doc.add_paragraph()
        r1 = p1.add_run(customer_text)
        r1.bold = True
        p2 = doc.add_paragraph()
        r2 = p2.add_run(discount_text)
        r2.bold = True
        return
    new_p = OxmlElement("w:p")
    table._tbl.addnext(new_p)
    paragraph = Paragraph(new_p, doc.element.body)
    run1 = paragraph.add_run(customer_text)
    run1.bold = True

    new_p2 = OxmlElement("w:p")
    paragraph._p.addnext(new_p2)
    paragraph2 = Paragraph(new_p2, doc.element.body)
    run2 = paragraph2.add_run(discount_text)
    run2.bold = True


def process(input_bytes: bytes, discount_percent: float) -> bytes:
    doc = Document(io.BytesIO(input_bytes))
    customer_id = extract_customer_id(doc) or "UNKNOWN"
    apply_image_to_items(doc)
    append_discount_note(doc, customer_id, discount_percent)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
